[CmdletBinding()]
param(
    [string]$Configuration = "Release",
    [ValidateSet("x64", "x86")]
    [string]$Platform = "x64"
)

$ErrorActionPreference = "Stop"
$root = Split-Path -Parent $PSScriptRoot
$project = Join-Path $root "src-windows\VisualTeX.VstoFlowAcceptance\VisualTeX.VstoFlowAcceptance.csproj"

$dotnet = @(
    (Join-Path $env:LOCALAPPDATA "Microsoft\dotnet\dotnet.exe"),
    (Join-Path $env:USERPROFILE ".dotnet\dotnet.exe"),
    (Join-Path $env:ProgramFiles "dotnet\dotnet.exe")
) | Where-Object { Test-Path $_ } | Select-Object -First 1
if (-not $dotnet) { throw ".NET 8 SDK is required." }
$dotnetVersion = & $dotnet --version
if ($LASTEXITCODE -ne 0 -or -not $dotnetVersion.StartsWith("8.")) {
    throw ".NET 8 SDK is required; resolved dotnet reports '$dotnetVersion'."
}

$vswhere = Join-Path ${env:ProgramFiles(x86)} "Microsoft Visual Studio\Installer\vswhere.exe"
$installationPath = & $vswhere -latest -version "[17.0,19.0)" -products * -requires Microsoft.Component.MSBuild -property installationPath
if (-not $installationPath) { throw "Visual Studio MSBuild was not found." }
$msbuild = Join-Path $installationPath "MSBuild\Current\Bin\amd64\MSBuild.exe"
if (-not (Test-Path $msbuild)) { throw "MSBuild was not found: $msbuild" }

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

$runtimeIdentifier = "win7-$Platform"
& $dotnet restore $project --ignore-failed-sources -p:Platform=$Platform -r $runtimeIdentifier
if ($LASTEXITCODE -ne 0) { throw "NuGet restore failed." }
if (-not (Test-Path $referenceRoot)) { throw ".NET Framework 4.8 reference assemblies are missing after restore." }
& $msbuild $project /m /p:Configuration=$Configuration /p:Platform=$Platform /p:TargetFrameworkRootPath=$referenceRoot
if ($LASTEXITCODE -ne 0) { throw "VSTO flow acceptance build failed." }
