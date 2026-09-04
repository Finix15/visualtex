[CmdletBinding()]
param(
    [string]$InstallerPath,
    [switch]$IUnderstandThisReplacesAndUninstallsVisualTeX
)

$ErrorActionPreference = "Stop"
Set-StrictMode -Version Latest

if (-not $IUnderstandThisReplacesAndUninstallsVisualTeX) {
    throw "This opt-in acceptance replaces the installed VisualTeX, verifies Office, and then uninstalls it. Re-run with -IUnderstandThisReplacesAndUninstallsVisualTeX."
}
if (Get-Process WINWORD, POWERPNT -ErrorAction SilentlyContinue) {
    throw "Word or PowerPoint is running. Save your documents and close Office before acceptance."
}
if (Get-Process visualtex -ErrorAction SilentlyContinue) {
    throw "VisualTeX is running. Close it before installed-release acceptance."
}

$root = Split-Path -Parent $PSScriptRoot
if ([string]::IsNullOrWhiteSpace($InstallerPath)) {
    $InstallerPath = Join-Path $root "src-tauri\target\release\bundle\nsis\VisualTeX_1.2.7_x64-setup.exe"
}
$InstallerPath = (Resolve-Path -LiteralPath $InstallerPath).Path
$uninstallKey = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall\VisualTeX"
$integrationKey = "HKCU:\Software\VisualTeX\OfficeIntegration"

function Get-StoreSnapshot([string]$StoreName) {
    $result = @{}
    foreach ($certificate in Get-ChildItem "Cert:\CurrentUser\$StoreName" -ErrorAction SilentlyContinue) {
        if ([string]$certificate.Subject -match '(^|,\s*)CN=VisualTeX Local Office Companion(,|$)') { continue }
        $result[[string]$certificate.Thumbprint] = [string]$certificate.Subject
    }
    return $result
}

function Assert-UnrelatedCertificatesRemain([string]$StoreName, [hashtable]$Before) {
    $after = @{}
    foreach ($certificate in Get-ChildItem "Cert:\CurrentUser\$StoreName" -ErrorAction SilentlyContinue) {
        $after[[string]$certificate.Thumbprint] = [string]$certificate.Subject
    }
    foreach ($thumbprint in $Before.Keys) {
        if (-not $after.ContainsKey($thumbprint)) {
            throw "Unrelated certificate $thumbprint ($($Before[$thumbprint])) disappeared from CurrentUser\\$StoreName."
        }
    }
}

$rootBefore = Get-StoreSnapshot "Root"
$trustedPeopleBefore = Get-StoreSnapshot "TrustedPeople"
$installedRoot = $null
$thumbprint = $null
$report = $null

try {
    $install = Start-Process -FilePath $InstallerPath `
        -ArgumentList @("/S", "/VISUALTEXOFFICE=native", "/VISUALTEXOCR=none") `
        -PassThru -Wait
    if ($install.ExitCode -ne 0) { throw "Installer failed with exit code $($install.ExitCode)." }

    $uninstallState = Get-ItemProperty -LiteralPath $uninstallKey -ErrorAction Stop
    $installedRoot = ([string]$uninstallState.InstallLocation).Trim().Trim('"')
    if ([string]::IsNullOrWhiteSpace($installedRoot)) { throw "Installed VisualTeX location is not registered." }
    $runtimeScript = Join-Path $installedRoot "scripts\test_windows_office_runtime.ps1"
    $installedExe = Join-Path $installedRoot "visualtex.exe"
    if (-not (Test-Path -LiteralPath $installedExe)) { $installedExe = Join-Path $installedRoot "VisualTeX.exe" }
    foreach ($forbiddenX86Payload in @(
        "windows-office\VisualTeX-WindowsOffice-VSTO-x86.msi",
        "windows-office\VisualTeX-WindowsOffice-VSTO-x86.sha256.json"
    )) {
        $forbiddenPath = Join-Path $installedRoot $forbiddenX86Payload
        if (Test-Path -LiteralPath $forbiddenPath) {
            throw "The x64-only VisualTeX installer wrote a forbidden Office x86 payload: $forbiddenPath"
        }
    }

    $integration = Get-ItemProperty -LiteralPath $integrationKey -ErrorAction Stop
    $thumbprint = ([string]$integration.CertificateThumbprint -replace '[^0-9A-Fa-f]', '').ToUpperInvariant()
    if ($thumbprint -notmatch '^[0-9A-F]{40}$') { throw "Installed certificate thumbprint is invalid: $thumbprint" }
    $rootCertificate = Get-Item -LiteralPath "Cert:\CurrentUser\Root\$thumbprint" -ErrorAction Stop
    if ([string]$rootCertificate.Subject -notmatch '(^|,\s*)CN=VisualTeX Local Office Companion(,|$)') {
        throw "Installed certificate has the wrong subject: $($rootCertificate.Subject)"
    }

    $report = Join-Path $env:TEMP "visualtex-personal-office-$([guid]::NewGuid().ToString('N')).json"
    & powershell.exe -NoProfile -ExecutionPolicy Bypass -File $runtimeScript `
        -OfficePlatform x64 -VisualTeXPath $installedExe -ReportPath $report
    if ($LASTEXITCODE -ne 0) { throw "Installed Office runtime verification failed with exit code $LASTEXITCODE." }

    $uninstaller = Join-Path $installedRoot "uninstall.exe"
    $uninstall = Start-Process -FilePath $uninstaller -ArgumentList "/S" -PassThru -Wait
    if ($uninstall.ExitCode -ne 0) { throw "VisualTeX uninstaller failed with exit code $($uninstall.ExitCode)." }

    $deadline = [DateTimeOffset]::UtcNow.AddSeconds(30)
    while ((Test-Path -LiteralPath $uninstaller) -and [DateTimeOffset]::UtcNow -lt $deadline) {
        Start-Sleep -Milliseconds 250
    }
    foreach ($storeName in @("Root", "TrustedPeople")) {
        if (Test-Path -LiteralPath "Cert:\CurrentUser\$storeName\$thumbprint") {
            throw "VisualTeX certificate $thumbprint remains in CurrentUser\\$storeName after uninstall."
        }
    }
    Assert-UnrelatedCertificatesRemain "Root" $rootBefore
    Assert-UnrelatedCertificatesRemain "TrustedPeople" $trustedPeopleBefore
    if (Get-Process WINWORD, POWERPNT -ErrorAction SilentlyContinue) {
        throw "Office process remains after installed acceptance."
    }
    Write-Host "Personal installed release acceptance passed: install, hidden Office verification, uninstall, and exact certificate cleanup."
} finally {
    if (-not [string]::IsNullOrWhiteSpace($report)) {
        Remove-Item -LiteralPath $report -Force -ErrorAction SilentlyContinue
    }
}
