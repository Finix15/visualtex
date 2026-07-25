[CmdletBinding()]
param(
    [string]$ExpectedAppVersion = "1.2.2",
    [string]$ExpectedOfficeMsiVersion = "1.0.35.0"
)

$ErrorActionPreference = "Stop"
$root = Split-Path -Parent $PSScriptRoot
$installerPath = Join-Path $root "src-tauri\target\release\bundle\nsis\VisualTeX_${ExpectedAppVersion}_x64-setup.exe"
$resourceX64 = Join-Path $root "src-tauri\resources\windows-office\VisualTeX-WindowsOffice-VSTO-x64.msi"
$resourceX86 = Join-Path $root "src-tauri\resources\windows-office\VisualTeX-WindowsOffice-VSTO-x86.msi"
$manifestX64 = Join-Path $root "src-tauri\resources\windows-office\VisualTeX-WindowsOffice-VSTO-x64.sha256.json"
$manifestX86 = Join-Path $root "src-tauri\resources\windows-office\VisualTeX-WindowsOffice-VSTO-x86.sha256.json"
$vstoRuntime = Join-Path $root "src-tauri\resources\windows-office\vstor_redist.exe"
$vstoRuntimeManifest = Join-Path $root "src-tauri\resources\windows-office\vstor_redist.sha256.json"
$buildX64 = Join-Path $root "src-windows\VisualTeX.WindowsOffice.Installer\bin\x64\Release\VisualTeX-WindowsOffice-VSTO-x64.msi"
$buildX86 = Join-Path $root "src-windows\VisualTeX.WindowsOffice.Installer\bin\x86\Release\VisualTeX-WindowsOffice-VSTO-x86.msi"
$paths = @($installerPath, $resourceX64, $resourceX86, $manifestX64, $manifestX86, $vstoRuntime, $vstoRuntimeManifest, $buildX64, $buildX86)

foreach ($path in $paths) {
    if (-not (Test-Path -LiteralPath $path)) { throw "Release artifact is missing: $path" }
    $item = Get-Item -LiteralPath $path
    $hash = Get-FileHash -LiteralPath $path -Algorithm SHA256
    Write-Host ("{0} | {1} bytes | SHA256 {2}" -f $item.FullName, $item.Length, $hash.Hash)
}

$installer = New-Object -ComObject WindowsInstaller.Installer
foreach ($path in @($resourceX64, $resourceX86, $buildX64, $buildX86)) {
    $database = $installer.OpenDatabase($path, 0)
    $view = $database.OpenView("SELECT `Value` FROM `Property` WHERE `Property`='ProductVersion'")
    $view.Execute()
    $record = $view.Fetch()
    $version = $record.StringData(1)
    Write-Host ("{0} | ProductVersion {1}" -f $path, $version)
    if ($version -ne $ExpectedOfficeMsiVersion) {
        throw "Unexpected Office MSI version in $path. Expected $ExpectedOfficeMsiVersion, actual $version."
    }
}

if ((Get-FileHash $resourceX64 -Algorithm SHA256).Hash -ne (Get-FileHash $buildX64 -Algorithm SHA256).Hash) {
    throw "The x64 Office MSI bundled by Tauri is not the current x64 build."
}
if ((Get-FileHash $resourceX86 -Algorithm SHA256).Hash -ne (Get-FileHash $buildX86 -Algorithm SHA256).Hash) {
    throw "The x86 Office MSI bundled by Tauri is not the current x86 build."
}

$runtimeManifest = Get-Content -LiteralPath $vstoRuntimeManifest -Raw | ConvertFrom-Json
$runtimeHash = (Get-FileHash -LiteralPath $vstoRuntime -Algorithm SHA256).Hash
if ([string]$runtimeManifest.package.sha256 -ne $runtimeHash) {
    throw "The bundled Microsoft VSTO Runtime hash manifest does not match $vstoRuntime."
}
if ($runtimeHash -ne "CFE1A40BBE4A50022DB2164ABDB0154984E2CECB761A23CDC81CB5754F6E0A18") {
    throw "The bundled Microsoft VSTO Runtime is not the pinned 10.0.60917.00 package."
}
$runtimeSignature = Get-AuthenticodeSignature -FilePath $vstoRuntime
if ($runtimeSignature.Status -ne [System.Management.Automation.SignatureStatus]::Valid -or
    $null -eq $runtimeSignature.SignerCertificate -or
    $runtimeSignature.SignerCertificate.Subject -notmatch "Microsoft Corporation") {
    throw "The bundled Microsoft VSTO Runtime has an invalid or unexpected Authenticode signature."
}
$runtimeVersion = (Get-Item -LiteralPath $vstoRuntime).VersionInfo
if ([string]$runtimeVersion.ProductVersion -ne "10.0.60917.00" -or
    [string]$runtimeVersion.CompanyName -ne "Microsoft Corporation") {
    throw "Unexpected bundled Microsoft VSTO Runtime metadata: Version=$($runtimeVersion.ProductVersion); Company=$($runtimeVersion.CompanyName)."
}

foreach ($entry in @(
    @{ Msi = $resourceX64; Manifest = $manifestX64 },
    @{ Msi = $resourceX86; Manifest = $manifestX86 }
)) {
    $manifest = Get-Content -LiteralPath $entry.Manifest -Raw | ConvertFrom-Json
    $actualHash = (Get-FileHash -LiteralPath $entry.Msi -Algorithm SHA256).Hash
    if ($manifest.package.sha256 -ne $actualHash) {
        throw "Office MSI hash manifest does not match $($entry.Msi)."
    }
}

Write-Host "VisualTeX Windows release artifacts passed static verification."
