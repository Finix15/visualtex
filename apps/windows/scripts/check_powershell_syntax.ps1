[CmdletBinding()]
param(
    [string[]]$Paths = @(
        "scripts/ensure_windows_office_certificate.ps1",
        "scripts/patch_generated_nsis.ps1",
        "scripts/test_windows_installed_release.ps1",
        "scripts/test_windows_office_connection_readonly.ps1",
        "scripts/verify_windows_release_artifacts.ps1"
    )
)

$ErrorActionPreference = "Stop"
Set-StrictMode -Version Latest

foreach ($path in $Paths) {
    $resolved = (Resolve-Path -LiteralPath $path).Path
    $tokens = $null
    $errors = $null
    [System.Management.Automation.Language.Parser]::ParseFile(
        $resolved,
        [ref]$tokens,
        [ref]$errors
    ) | Out-Null
    if ($errors.Count -gt 0) {
        $messages = $errors | ForEach-Object {
            "{0}:{1}:{2}: {3}" -f $resolved, $_.Extent.StartLineNumber, $_.Extent.StartColumnNumber, $_.Message
        }
        throw ($messages -join [Environment]::NewLine)
    }
    Write-Host "Parsed $resolved"
}

Write-Host "PowerShell syntax verification passed."
