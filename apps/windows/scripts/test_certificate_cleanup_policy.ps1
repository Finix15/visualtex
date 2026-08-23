[CmdletBinding()]
param()

$ErrorActionPreference = "Stop"
$testKey = "HKCU:\Software\VisualTeX\Tests\CertificateCleanup-$([guid]::NewGuid().ToString('N'))"
$logPath = Join-Path $env:TEMP "visualtex-certificate-cleanup-policy-$([guid]::NewGuid().ToString('N')).log"
$scriptPath = Join-Path $PSScriptRoot "remove_windows_office_certificate.ps1"
$uninstallPath = Join-Path $PSScriptRoot "uninstall_windows_vsto.ps1"

try {
    & powershell.exe -NoProfile -NonInteractive -ExecutionPolicy Bypass -File $scriptPath `
        -ModeKey $testKey -LogPath $logPath -SimulateFailure
    if ($LASTEXITCODE -ne 1) {
        throw "Simulated certificate cleanup returned $LASTEXITCODE instead of 1."
    }
    $state = Get-ItemProperty -LiteralPath $testKey -ErrorAction Stop
    if ([int]$state.CertificateCleanupPending -ne 1 -or
        [string]::IsNullOrWhiteSpace([string]$state.CertificateCleanupError) -or
        -not [string]::Equals([string]$state.CertificateCleanupLogPath, $logPath, [StringComparison]::OrdinalIgnoreCase)) {
        throw "Certificate cleanup failure state was not persisted correctly."
    }

    $uninstallSource = Get-Content -LiteralPath $uninstallPath -Raw
    if ($uninstallSource -notmatch 'certificate cleanup failed\. Uninstall will continue' -or
        $uninstallSource -match 'throw "VisualTeX Office certificate cleanup failed') {
        throw "Office uninstall does not preserve the non-blocking certificate cleanup policy."
    }
    Write-Host "Certificate cleanup failure policy passed without changing the real certificate stores."
} finally {
    Remove-Item -LiteralPath $testKey -Recurse -Force -ErrorAction SilentlyContinue
    Remove-Item -LiteralPath $logPath -Force -ErrorAction SilentlyContinue
    $testsRoot = "HKCU:\Software\VisualTeX\Tests"
    if (Test-Path -LiteralPath $testsRoot) {
        $children = @(Get-ChildItem -LiteralPath $testsRoot -ErrorAction SilentlyContinue)
        if ($children.Count -eq 0) { Remove-Item -LiteralPath $testsRoot -Force -ErrorAction SilentlyContinue }
    }
}
