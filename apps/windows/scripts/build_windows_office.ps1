[CmdletBinding()]
param(
    [string]$Configuration = "Release",
    [switch]$SkipTests,
    [ValidateSet("x64", "x86")]
    [string[]]$Platforms = @("x64")
)

$ErrorActionPreference = "Stop"
Import-Module Microsoft.PowerShell.Utility -Force -ErrorAction Stop
$root = Split-Path -Parent $PSScriptRoot
$solution = Join-Path $root "src-windows\VisualTeX.WindowsOffice.sln"
$tests = Join-Path $root "src-windows\VisualTeX.WindowsOffice.Tests\VisualTeX.WindowsOffice.Tests.csproj"
$wordProject = Join-Path $root "src-windows\VisualTeX.WordVsto\VisualTeX.WordVsto.csproj"
$powerPointProject = Join-Path $root "src-windows\VisualTeX.PowerPointVsto\VisualTeX.PowerPointVsto.csproj"
$installerProject = Join-Path $root "src-windows\VisualTeX.WindowsOffice.Installer\VisualTeX.WindowsOffice.Installer.wixproj"
$oleServerProject = Join-Path $root "src-windows\VisualTeX.FormulaOleServer\VisualTeX.FormulaOleServer.vcxproj"
$resourceRoot = Join-Path $root "src-tauri\resources\windows-office"
$oleBuildRoot = Join-Path $root "src-windows\artifacts\formula-ole-server"

function Stop-BuildOleServerProcesses {
    $normalizedRoot = [IO.Path]::GetFullPath($oleBuildRoot).TrimEnd('\') + '\'
    foreach ($process in Get-CimInstance Win32_Process -Filter "Name='VisualTeX.FormulaOleServer.exe'" -ErrorAction SilentlyContinue) {
        $path = [string]$process.ExecutablePath
        if ([string]::IsNullOrWhiteSpace($path)) { continue }
        $normalizedPath = [IO.Path]::GetFullPath($path)
        if (-not $normalizedPath.StartsWith($normalizedRoot, [StringComparison]::OrdinalIgnoreCase)) { continue }
        Stop-Process -Id ([int]$process.ProcessId) -Force -ErrorAction SilentlyContinue
        Write-Host "Stopped acceptance-owned OLE Server PID=$($process.ProcessId) before rebuild: $normalizedPath"
    }
    Start-Sleep -Milliseconds 300
}

$dotnet = $null
$dotnetVersion = $null
foreach ($candidate in @(
    (Join-Path $env:LOCALAPPDATA "Microsoft\dotnet\dotnet.exe"),
    (Join-Path $env:USERPROFILE ".dotnet\dotnet.exe"),
    (Join-Path $env:ProgramFiles "dotnet\dotnet.exe")
)) {
    if (-not (Test-Path $candidate)) { continue }
    $candidateVersion = @(& $candidate --list-sdks) |
        ForEach-Object { ($_ -split ' ')[0] } |
        Where-Object { $_.StartsWith("8.") } |
        Sort-Object { [version]$_ } -Descending |
        Select-Object -First 1
    if ($candidateVersion) {
        $dotnet = $candidate
        $dotnetVersion = $candidateVersion
        break
    }
}
if (-not $dotnet) {
    $dotnet = Get-Command dotnet.exe -ErrorAction SilentlyContinue | Select-Object -First 1 -ExpandProperty Source
}
if (-not $dotnet) { throw ".NET 8 SDK is required to build and test Windows Office integration." }
if (-not $dotnetVersion) { $dotnetVersion = & $dotnet --version }
if ($LASTEXITCODE -ne 0 -or -not $dotnetVersion.StartsWith("8.")) {
    throw ".NET 8 SDK is required; resolved dotnet reports '$dotnetVersion'."
}
Write-Host "Using dotnet $dotnetVersion from $dotnet"

if (-not $SkipTests) {
    Push-Location $root
    try { & $dotnet test $tests --configuration $Configuration }
    finally { Pop-Location }
    if ($LASTEXITCODE -ne 0) { throw "Windows Office tests failed." }
    & (Join-Path $PSScriptRoot "test_windows_formula_ole_server.ps1") -Configuration $Configuration
    if ($LASTEXITCODE -ne 0) { throw "Native Formula OLE tests failed." }
    Stop-BuildOleServerProcesses
}

& (Join-Path $PSScriptRoot "build_windows_ole_bridge.ps1") -Configuration $Configuration

$msbuild = Get-Command msbuild.exe -ErrorAction SilentlyContinue | Select-Object -First 1 -ExpandProperty Source
if (-not $msbuild) {
    $vswhere = Join-Path ${env:ProgramFiles(x86)} "Microsoft Visual Studio\Installer\vswhere.exe"
    if (Test-Path $vswhere) {
        $installationPath = & $vswhere -latest -version "[17.0,19.0)" -products * `
            -requires Microsoft.Component.MSBuild Microsoft.VisualStudio.Component.VC.Tools.x86.x64 Microsoft.VisualStudio.Component.VC.ATL `
            -property installationPath
        if ($installationPath) {
            $candidate = Join-Path $installationPath "MSBuild\Current\Bin\amd64\MSBuild.exe"
            if (Test-Path $candidate) { $msbuild = $candidate }
        }
    }
}
if (-not $msbuild) {
    throw "Visual Studio Build Tools with MSBuild, Visual C++ x86/x64 and ATL is required for Word/PowerPoint native add-ins and WiX MSI."
}
$sdkRoot = Join-Path (Split-Path -Parent $dotnet) "sdk"
$sdkPath = Join-Path $sdkRoot $dotnetVersion
$sdk = Get-Item -LiteralPath $sdkPath -ErrorAction SilentlyContinue
if (-not $sdk) { throw ".NET SDK directory matching dotnet $dotnetVersion was not found at $sdkPath." }
$env:DOTNET_ROOT = Split-Path -Parent $dotnet
$env:DOTNET_HOST_PATH = $dotnet
$env:DOTNET_MSBUILD_SDK_RESOLVER_CLI_DIR = Split-Path -Parent $dotnet
$env:MSBuildSDKsPath = Join-Path $sdk.FullName "Sdks"
$env:MSBuildEnableWorkloadResolver = "false"
$referenceRoot = Join-Path $env:USERPROFILE ".nuget\packages\microsoft.netframework.referenceassemblies.net48\1.0.3\build"

New-Item $resourceRoot -ItemType Directory -Force | Out-Null

$architectures = @(
    [ordered]@{ PackagePlatform = "x64"; OlePlatform = "x64" },
    [ordered]@{ PackagePlatform = "x86"; OlePlatform = "Win32" }
) | Where-Object { $Platforms -contains $_.PackagePlatform }

foreach ($architecture in $architectures) {
    $packagePlatform = $architecture.PackagePlatform
    $olePlatform = $architecture.OlePlatform
    Write-Host "Building native Office integration for $packagePlatform Office."
    Stop-BuildOleServerProcesses

    & $msbuild $oleServerProject /m /p:Configuration=$Configuration /p:Platform=$olePlatform
    if ($LASTEXITCODE -ne 0) {
        throw "$packagePlatform native Formula OLE LocalServer build failed."
    }

    $runtimeIdentifier = "win7-$packagePlatform"
    foreach ($project in @($wordProject, $powerPointProject)) {
        Push-Location $root
        try {
            & $dotnet restore $project --ignore-failed-sources -p:Platform=$packagePlatform -p:RestoreBuildInParallel=false -r $runtimeIdentifier
        }
        finally {
            Pop-Location
        }
        if ($LASTEXITCODE -ne 0) { throw "NuGet restore failed: $project" }
        if (-not (Test-Path $referenceRoot)) {
            throw ".NET Framework 4.8 reference assemblies package is missing after restoring $project."
        }
        & $msbuild $project /m:1 /p:Configuration=$Configuration /p:Platform=$packagePlatform /p:TargetFrameworkRootPath=$referenceRoot /p:UseSharedCompilation=false
        if ($LASTEXITCODE -ne 0) { throw "$packagePlatform VSTO build failed: $project" }
    }

    & (Join-Path $PSScriptRoot "test_windows_vsto_ribbon_dispatch.ps1") `
        -Configuration $Configuration `
        -Platform $packagePlatform
    & (Join-Path $PSScriptRoot "test_windows_vsto_dependency_loading.ps1") `
        -Configuration $Configuration `
        -Platform $packagePlatform

    & $msbuild $installerProject `
        /m `
        /p:Configuration=$Configuration `
        /p:Platform=$packagePlatform `
        /p:BuildProjectReferences=false `
        /p:SuppressValidation=true
    if ($LASTEXITCODE -ne 0) { throw "$packagePlatform Windows Office MSI build failed." }

    $packageFileName = "VisualTeX-WindowsOffice-VSTO-$packagePlatform.msi"
    $msi = Join-Path $root "src-windows\VisualTeX.WindowsOffice.Installer\bin\$packagePlatform\$Configuration\$packageFileName"
    if (-not (Test-Path $msi)) { throw "$packagePlatform VisualTeX VSTO MSI was not produced: $msi" }
    Copy-Item $msi (Join-Path $resourceRoot $packageFileName) -Force

    $wordOutput = Join-Path $root "src-windows\VisualTeX.WordVsto\bin\$packagePlatform\$Configuration\net48\VisualTeX.WordVsto.dll"
    $powerPointOutput = Join-Path $root "src-windows\VisualTeX.PowerPointVsto\bin\$packagePlatform\$Configuration\net48\VisualTeX.PowerPointVsto.dll"
    $oleServerOutput = Join-Path $root "src-windows\artifacts\formula-ole-server\$olePlatform\$Configuration\VisualTeX.FormulaOleServer.exe"
    if (-not (Test-Path $wordOutput) -or -not (Test-Path $powerPointOutput) -or -not (Test-Path $oleServerOutput)) {
        throw "The $packagePlatform VSTO and native OLE outputs required for SHA-256 verification are missing."
    }
    $dependencyFileNames = @(
        "VisualTeX.WindowsOffice.Contracts.dll",
        "VisualTeX.MathTypeConversion.dll",
        "System.Text.Json.dll",
        "System.Text.Encodings.Web.dll",
        "Microsoft.Bcl.AsyncInterfaces.dll",
        "System.Memory.dll",
        "System.Buffers.dll",
        "System.Numerics.Vectors.dll",
        "System.ValueTuple.dll",
        "System.Runtime.CompilerServices.Unsafe.dll",
        "System.Threading.Tasks.Extensions.dll"
    )
    $wordOutputDirectory = Split-Path -Parent $wordOutput
    $dependencyEntries = @()
    foreach ($dependencyFileName in $dependencyFileNames) {
        $dependencyPath = Join-Path $wordOutputDirectory $dependencyFileName
        if (-not (Test-Path $dependencyPath)) {
            throw "The $packagePlatform VSTO dependency required for packaging is missing: $dependencyPath"
        }
        $dependencyEntries += [ordered]@{
            file = $dependencyFileName
            sha256 = (Get-FileHash $dependencyPath -Algorithm SHA256).Hash
        }
    }

    $hashManifest = [ordered]@{
        version = "1.0.43.0"
        architecture = $packagePlatform
        package = [ordered]@{
            file = $packageFileName
            sha256 = (Get-FileHash $msi -Algorithm SHA256).Hash
        }
        word = [ordered]@{
            file = "VisualTeX.WordVsto.dll"
            sha256 = (Get-FileHash $wordOutput -Algorithm SHA256).Hash
        }
        powerPoint = [ordered]@{
            file = "VisualTeX.PowerPointVsto.dll"
            sha256 = (Get-FileHash $powerPointOutput -Algorithm SHA256).Hash
        }
        formulaOleServer = [ordered]@{
            file = "VisualTeX.FormulaOleServer.exe"
            sha256 = (Get-FileHash $oleServerOutput -Algorithm SHA256).Hash
        }
        dependencies = $dependencyEntries
        mathTypeRuntime = [ordered]@{
            protocolVersion = 1
            files = @(Get-ChildItem (Join-Path $root "src-windows\artifacts\mathtype-runtime\windows-x64") -Recurse -File | ForEach-Object {
                [ordered]@{
                    path = $_.FullName.Substring((Join-Path $root "src-windows\artifacts\mathtype-runtime\windows-x64").Length + 1).Replace('\', '/')
                    sha256 = (Get-FileHash $_.FullName -Algorithm SHA256).Hash
                }
            })
        }
    }
    $manifestFileName = "VisualTeX-WindowsOffice-VSTO-$packagePlatform.sha256.json"
    $hashManifest | ConvertTo-Json -Depth 4 | Set-Content `
        (Join-Path $resourceRoot $manifestFileName) `
        -Encoding UTF8
}

Write-Host "Windows Office package(s) ready for: $($Platforms -join ', ')."
