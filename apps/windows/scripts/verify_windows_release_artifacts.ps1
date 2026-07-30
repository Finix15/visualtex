[CmdletBinding()]
param(
    [string]$ExpectedAppVersion = "1.2.3",
    [string]$ExpectedOfficeMsiVersion = "1.0.40.0"
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

function Assert-MsiComponentBitness {
    param(
        [object]$Installer,
        [string]$Path,
        [bool]$Expected64Bit
    )
    $requiredComponents = @(
        "VstoFiles",
        "NativeOleServerFiles",
        "NativeOleRegistration",
        "WordComRegistration",
        "PowerPointComRegistration",
        "WordOfficeRegistration",
        "PowerPointOfficeRegistration",
        "ModeRegistration"
    )
    $database = $null
    $view = $null
    try {
        $database = $Installer.OpenDatabase($Path, 0)
        $view = $database.OpenView('SELECT `Component`, `Attributes` FROM `Component`')
        $view.Execute()
        $found = @{}
        while ($true) {
            $record = $view.Fetch()
            if ($null -eq $record) { break }
            try {
                $name = [string]$record.StringData(1)
                $attributes = [int]$record.IntegerData(2)
                $found[$name] = $attributes
            } finally {
                if ([Runtime.InteropServices.Marshal]::IsComObject($record)) {
                    [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($record)
                }
            }
        }
        foreach ($component in $requiredComponents) {
            if (-not $found.ContainsKey($component)) {
                throw "Office MSI is missing required component '$component': $Path"
            }
            $is64Bit = ([int]$found[$component] -band 256) -ne 0
            if ($is64Bit -ne $Expected64Bit) {
                throw "Office MSI component '$component' has the wrong registry/file bitness in $Path. Expected64Bit=$Expected64Bit; Attributes=$($found[$component])."
            }
        }
        Write-Host ("{0} | Component bitness {1} verified for {2} components" -f $Path, $(if ($Expected64Bit) { "x64" } else { "x86" }), $requiredComponents.Count)
    } finally {
        if ($null -ne $view -and [Runtime.InteropServices.Marshal]::IsComObject($view)) {
            [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($view)
        }
        if ($null -ne $database -and [Runtime.InteropServices.Marshal]::IsComObject($database)) {
            [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($database)
        }
    }
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
    $scopeView = $database.OpenView("SELECT `Value` FROM `Property` WHERE `Property`='ALLUSERS'")
    $scopeView.Execute()
    $scopeRecord = $scopeView.Fetch()
    $allUsers = if ($null -eq $scopeRecord) { "" } else { [string]$scopeRecord.StringData(1) }
    if ($allUsers -ne "1") {
        throw "Office MSI is not machine-wide: $path; ALLUSERS='$allUsers'."
    }
    Write-Host ("{0} | ALLUSERS=1 machine-wide scope verified" -f $path)
    if ($null -ne $scopeRecord -and [Runtime.InteropServices.Marshal]::IsComObject($scopeRecord)) {
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($scopeRecord)
    }
    if ([Runtime.InteropServices.Marshal]::IsComObject($scopeView)) {
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($scopeView)
    }
}
Assert-MsiComponentBitness $installer $resourceX64 $true
Assert-MsiComponentBitness $installer $buildX64 $true
Assert-MsiComponentBitness $installer $resourceX86 $false
Assert-MsiComponentBitness $installer $buildX86 $false

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
