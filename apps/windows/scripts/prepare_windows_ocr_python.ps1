param(
    [switch]$Force
)

$ErrorActionPreference = "Stop"
Set-StrictMode -Version Latest

$PythonVersion = "3.12.10"
$PythonArchiveName = "python-$PythonVersion-embed-amd64.zip"
$PythonArchiveUrl = "https://www.python.org/ftp/python/$PythonVersion/$PythonArchiveName"
$PythonArchiveSha256 = "4ACBED6DD1C744B0376E3B1CF57CE906F9DC9E95E68824584C8099A63025A3C3"
$PipVersion = "25.1.1"
$PipWheelName = "pip-$PipVersion-py3-none-any.whl"
$PipWheelUrl = "https://files.pythonhosted.org/packages/29/a2/d40fb2460e883eca5199c62cfc2463fd261f760556ae6290f88488c362c0/$PipWheelName"
$PipWheelSha256 = "2913A38A2ABF4EA6B64AB507BD9E967F3B53DC1EDE74B01B0931E1CE548751AF"

$ProjectRoot = (Resolve-Path (Join-Path $PSScriptRoot "..")).Path
$SiteCustomizeSource = Join-Path $ProjectRoot "src-tauri\ocr\private_sitecustomize.py"
$ResourceRoot = Join-Path $ProjectRoot "src-tauri\resources\ocr-python\windows-x64"
$OutputArchiveName = "visualtex-python-$PythonVersion-windows-x64.zip"
$OutputArchive = Join-Path $ResourceRoot $OutputArchiveName
$OutputManifest = Join-Path $ResourceRoot "manifest.json"
$CacheRoot = Join-Path $env:LOCALAPPDATA "VisualTeXBuildCache\ocr-python"
$PythonArchive = Join-Path $CacheRoot $PythonArchiveName
$PipWheel = Join-Path $CacheRoot $PipWheelName

function Get-Sha256([string]$Path) {
    return (Get-FileHash -LiteralPath $Path -Algorithm SHA256).Hash.ToUpperInvariant()
}

function Write-Utf8NoBom([string]$Path, [string]$Content) {
    $encoding = New-Object System.Text.UTF8Encoding($false)
    [System.IO.File]::WriteAllText($Path, $Content, $encoding)
}

function Get-VerifiedDownload(
    [string]$Url,
    [string]$Path,
    [string]$ExpectedSha256
) {
    if (Test-Path -LiteralPath $Path) {
        $actual = Get-Sha256 $Path
        if ($actual -eq $ExpectedSha256) {
            return
        }
        Remove-Item -LiteralPath $Path -Force
    }

    $directory = Split-Path -Parent $Path
    New-Item -ItemType Directory -Path $directory -Force | Out-Null
    $temporary = "$Path.download"
    Remove-Item -LiteralPath $temporary -Force -ErrorAction SilentlyContinue
    Write-Host "Downloading $Url"
    Invoke-WebRequest -UseBasicParsing -Uri $Url -OutFile $temporary
    $actual = Get-Sha256 $temporary
    if ($actual -ne $ExpectedSha256) {
        Remove-Item -LiteralPath $temporary -Force -ErrorAction SilentlyContinue
        throw "SHA-256 mismatch for $Url. Expected $ExpectedSha256, actual $actual"
    }
    Move-Item -LiteralPath $temporary -Destination $Path -Force
}

function Test-PreparedRuntime([string]$RuntimeRoot) {
    $python = Join-Path $RuntimeRoot "python.exe"
    if (-not (Test-Path -LiteralPath $python -PathType Leaf)) {
        throw "Prepared OCR Python is missing python.exe"
    }
    $probeScript = Join-Path $RuntimeRoot "_visualtex_runtime_probe.py"
    @'
import importlib.metadata as metadata
import json
import os
import platform
import site
import sys
user_site = site.getusersitepackages()
user_sites = [user_site] if isinstance(user_site, str) else list(user_site)
normalized = lambda value: os.path.normcase(os.path.abspath(value))
user_site_on_path = any(
    normalized(path) == normalized(candidate)
    for path in sys.path
    for candidate in user_sites
    if candidate
)
print(json.dumps({
    "pythonVersion": platform.python_version(),
    "pipVersion": metadata.version("pip"),
    "executable": sys.executable,
    "prefix": sys.prefix,
    "userSiteEnabled": bool(site.ENABLE_USER_SITE),
    "userSiteOnPath": user_site_on_path,
}))
'@ | Set-Content -LiteralPath $probeScript -Encoding UTF8
    try {
        $probe = & $python $probeScript
        if ($LASTEXITCODE -ne 0) {
            throw "Prepared OCR Python validation failed with exit code $LASTEXITCODE"
        }
    } finally {
        Remove-Item -LiteralPath $probeScript -Force -ErrorAction SilentlyContinue
    }
    $parsed = $probe | Select-Object -Last 1 | ConvertFrom-Json
    if ($parsed.pythonVersion -ne $PythonVersion) {
        throw "Prepared OCR Python version is $($parsed.pythonVersion), expected $PythonVersion"
    }
    if ($parsed.pipVersion -ne $PipVersion) {
        throw "Prepared OCR pip version is $($parsed.pipVersion), expected $PipVersion"
    }
    if ($parsed.userSiteEnabled -or $parsed.userSiteOnPath) {
        throw "Prepared OCR Python is not isolated from user site-packages"
    }
    $pipOutput = & $python -m pip --version
    if ($LASTEXITCODE -ne 0 -or -not ($pipOutput -match [regex]::Escape($RuntimeRoot))) {
        throw "python -m pip did not resolve inside the prepared private runtime: $pipOutput"
    }
}

New-Item -ItemType Directory -Path $ResourceRoot -Force | Out-Null
New-Item -ItemType Directory -Path $CacheRoot -Force | Out-Null
Get-VerifiedDownload $PythonArchiveUrl $PythonArchive $PythonArchiveSha256
Get-VerifiedDownload $PipWheelUrl $PipWheel $PipWheelSha256

$workRoot = Join-Path $env:TEMP ("visualtex-ocr-python-build-" + [guid]::NewGuid().ToString("N"))
$runtimeRoot = Join-Path $workRoot "runtime"
$validationRoot = Join-Path $workRoot "validation"
try {
    New-Item -ItemType Directory -Path $runtimeRoot -Force | Out-Null
    Expand-Archive -LiteralPath $PythonArchive -DestinationPath $runtimeRoot -Force

    $sitePackages = Join-Path $runtimeRoot "Lib\site-packages"
    New-Item -ItemType Directory -Path $sitePackages -Force | Out-Null
    $pipWheelZip = Join-Path $workRoot "pip-wheel.zip"
    Copy-Item -LiteralPath $PipWheel -Destination $pipWheelZip -Force
    Expand-Archive -LiteralPath $pipWheelZip -DestinationPath $sitePackages -Force

    $pthPath = Join-Path $runtimeRoot "python312._pth"
    @(
        "python312.zip",
        ".",
        "Lib\site-packages",
        "import site"
    ) | Set-Content -LiteralPath $pthPath -Encoding ASCII

    if (-not (Test-Path -LiteralPath $SiteCustomizeSource -PathType Leaf)) {
        throw "VisualTeX private Python sitecustomize source is missing: $SiteCustomizeSource"
    }
    $siteCustomize = Get-Content -LiteralPath $SiteCustomizeSource -Raw
    Write-Utf8NoBom (Join-Path $sitePackages "sitecustomize.py") $siteCustomize

    $runtimeMetadata = [ordered]@{
        schemaVersion = 1
        pythonVersion = $PythonVersion
        architecture = "x64"
        distribution = "python.org embeddable"
        pipVersion = $PipVersion
    }
    Write-Utf8NoBom (Join-Path $runtimeRoot "visualtex-python.json") ($runtimeMetadata | ConvertTo-Json -Depth 4)

    Test-PreparedRuntime $runtimeRoot

    # Build and validate a unique staged archive before replacing the bundled
    # runtime. Compress-Archive can observe a just-deleted destination on Windows
    # (for example while Defender still has the old ZIP open) and incorrectly
    # report that the file already exists. Staging also preserves the last known
    # good bundle if compression or validation fails.
    $stagedArchive = Join-Path $workRoot $OutputArchiveName
    Compress-Archive -Path (Join-Path $runtimeRoot "*") -DestinationPath $stagedArchive -CompressionLevel Optimal

    Remove-Item -LiteralPath $validationRoot -Recurse -Force -ErrorAction SilentlyContinue
    New-Item -ItemType Directory -Path $validationRoot -Force | Out-Null
    Expand-Archive -LiteralPath $stagedArchive -DestinationPath $validationRoot -Force
    Test-PreparedRuntime $validationRoot

    $archiveInstalled = $false
    for ($attempt = 1; $attempt -le 5; $attempt++) {
        try {
            if (Test-Path -LiteralPath $OutputArchive) {
                Remove-Item -LiteralPath $OutputArchive -Force -ErrorAction Stop
            }
            Move-Item -LiteralPath $stagedArchive -Destination $OutputArchive -Force -ErrorAction Stop
            $archiveInstalled = $true
            break
        } catch {
            if ($attempt -eq 5) {
                throw "Unable to replace bundled OCR Python after $attempt attempts: $($_.Exception.Message)"
            }
            Start-Sleep -Milliseconds (200 * $attempt)
        }
    }
    if (-not $archiveInstalled -or -not (Test-Path -LiteralPath $OutputArchive -PathType Leaf)) {
        throw "Bundled OCR Python archive was not installed: $OutputArchive"
    }

    $archiveInfo = Get-Item -LiteralPath $OutputArchive
    $manifest = [ordered]@{
        schemaVersion = 1
        platform = "windows"
        architecture = "x64"
        pythonVersion = $PythonVersion
        pipVersion = $PipVersion
        archive = [ordered]@{
            name = $OutputArchiveName
            size = $archiveInfo.Length
            sha256 = Get-Sha256 $OutputArchive
        }
        upstream = [ordered]@{
            pythonArchive = $PythonArchiveName
            pythonArchiveSha256 = $PythonArchiveSha256
            pipWheel = $PipWheelName
            pipWheelSha256 = $PipWheelSha256
        }
    }
    Write-Utf8NoBom $OutputManifest ($manifest | ConvertTo-Json -Depth 6)

    Write-Host "Prepared bundled OCR Python: $OutputArchive"
    Write-Host "SHA-256: $($manifest.archive.sha256)"
    Write-Host "Size: $($manifest.archive.size) bytes"
} finally {
    Remove-Item -LiteralPath $workRoot -Recurse -Force -ErrorAction SilentlyContinue
}
