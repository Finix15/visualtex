[CmdletBinding()]
param(
    [int]$TimeoutSeconds = 10,
    [string]$LogPath,
    [string]$ModeKey = "HKCU:\Software\VisualTeX\OfficeIntegration",
    [switch]$SimulateFailure
)

$ErrorActionPreference = "Stop"
$modeKey = $ModeKey
$logRoot = Join-Path $env:LOCALAPPDATA "VisualTeX\office\install-logs"
New-Item -Path $logRoot -ItemType Directory -Force | Out-Null
if ([string]::IsNullOrWhiteSpace($LogPath)) {
    $LogPath = Join-Path $logRoot ("certificate-remove-{0}.log" -f (Get-Date -Format "yyyyMMdd-HHmmss"))
}

function Write-CertificateLog([string]$Message) {
    $line = "[{0:O}] {1}" -f (Get-Date), $Message
    Write-Host $line
    Add-Content -LiteralPath $LogPath -Value $line -Encoding UTF8
}

function Test-CertificatePresent([string]$Thumbprint) {
    foreach ($storeName in @("Root", "TrustedPeople")) {
        if (Get-VisualTeXCertificate $storeName $Thumbprint) { return $true }
    }
    return $false
}

function Normalize-Thumbprint([string]$Value) {
    if ([string]::IsNullOrWhiteSpace($Value)) { return "" }
    return ($Value -replace '[^0-9A-Fa-f]', '').ToUpperInvariant()
}

function Get-VisualTeXCertificate([string]$StoreName, [string]$Thumbprint) {
    $normalized = Normalize-Thumbprint $Thumbprint
    if ($normalized -notmatch '^[0-9A-F]{40}$') { return $null }
    $certificate = Get-ChildItem "Cert:\CurrentUser\$StoreName" -ErrorAction SilentlyContinue |
        Where-Object { (Normalize-Thumbprint $_.Thumbprint) -eq $normalized } |
        Select-Object -First 1
    if ($null -eq $certificate) { return $null }
    if ([string]$certificate.Subject -notmatch '(^|,\s*)CN=VisualTeX Local Office Companion(,|$)') {
        throw "Refusing to remove certificate $normalized from CurrentUser\\$StoreName because its subject is '$($certificate.Subject)'."
    }
    return $certificate
}

$thumbprints = New-Object System.Collections.Generic.HashSet[string]([StringComparer]::OrdinalIgnoreCase)
$certificatePaths = New-Object System.Collections.Generic.List[string]
if (Test-Path -LiteralPath $modeKey) {
    $integration = Get-ItemProperty -LiteralPath $modeKey -ErrorAction SilentlyContinue
    $registeredThumbprint = Normalize-Thumbprint ([string]$integration.CertificateThumbprint)
    if ($registeredThumbprint) { [void]$thumbprints.Add($registeredThumbprint) }
    if (-not [string]::IsNullOrWhiteSpace([string]$integration.CertificatePath)) {
        [void]$certificatePaths.Add([string]$integration.CertificatePath)
    }
    if (-not [string]::IsNullOrWhiteSpace([string]$integration.AppDataRoot)) {
        [void]$certificatePaths.Add((Join-Path ([string]$integration.AppDataRoot) "office\localhost-cert.pem"))
    }
}

foreach ($certificatePath in @($certificatePaths)) {
    if (-not (Test-Path -LiteralPath $certificatePath -PathType Leaf)) { continue }
    try {
        $fileCertificate = [Security.Cryptography.X509Certificates.X509Certificate2]::new($certificatePath)
        try {
            if ([string]$fileCertificate.Subject -match '(^|,\s*)CN=VisualTeX Local Office Companion(,|$)') {
                [void]$thumbprints.Add((Normalize-Thumbprint $fileCertificate.Thumbprint))
            } else {
                Write-CertificateLog "Ignoring certificate file with unexpected subject: $certificatePath; Subject=$($fileCertificate.Subject)"
            }
        } finally { $fileCertificate.Dispose() }
    } catch {
        Write-CertificateLog "Unable to inspect certificate candidate ${certificatePath}: $($_.Exception.Message)"
    }
}

try {
    if ($SimulateFailure) {
        throw "Simulated VisualTeX certificate cleanup failure."
    }
    if ($thumbprints.Count -eq 0) {
        Write-CertificateLog "No VisualTeX certificate thumbprint is registered; nothing to remove."
    } else {
        foreach ($thumbprint in @($thumbprints)) {
            foreach ($storeName in @("Root", "TrustedPeople")) {
                $certificate = Get-VisualTeXCertificate $storeName $thumbprint
                if ($null -eq $certificate) {
                    Write-CertificateLog "Certificate $thumbprint is absent from CurrentUser\\$storeName."
                    continue
                }
                Write-CertificateLog "Removing exact VisualTeX certificate $thumbprint from CurrentUser\\$storeName."
                Remove-Item -LiteralPath "Cert:\CurrentUser\$storeName\$thumbprint" -Force -ErrorAction Stop
                if ($null -ne (Get-VisualTeXCertificate $storeName $thumbprint)) {
                    throw "Certificate $thumbprint remains in CurrentUser\\$storeName after removal."
                }
            }
        }
    }

    if (Test-Path -LiteralPath $modeKey) {
        Remove-ItemProperty -LiteralPath $modeKey -Name CertificateThumbprint -Force -ErrorAction SilentlyContinue
        Remove-ItemProperty -LiteralPath $modeKey -Name CertificatePath -Force -ErrorAction SilentlyContinue
        Remove-ItemProperty -LiteralPath $modeKey -Name CertificateCleanupPending -Force -ErrorAction SilentlyContinue
        Remove-ItemProperty -LiteralPath $modeKey -Name CertificateCleanupError -Force -ErrorAction SilentlyContinue
        Remove-ItemProperty -LiteralPath $modeKey -Name CertificateCleanupLogPath -Force -ErrorAction SilentlyContinue
    }
    Write-CertificateLog "VisualTeX current-user Office HTTPS certificate trust cleanup completed."
    exit 0
} catch {
    Write-CertificateLog ("FAILED: " + $_.Exception.ToString())
    New-Item -Path $modeKey -Force | Out-Null
    New-ItemProperty -LiteralPath $modeKey -Name CertificateCleanupPending -PropertyType DWord -Value 1 -Force | Out-Null
    New-ItemProperty -LiteralPath $modeKey -Name CertificateCleanupError -PropertyType String -Value $_.Exception.Message -Force | Out-Null
    New-ItemProperty -LiteralPath $modeKey -Name CertificateCleanupLogPath -PropertyType String -Value $LogPath -Force | Out-Null
    exit 1
}
