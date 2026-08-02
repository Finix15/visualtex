param(
    [string]$BaseUrl = "https://download.visualtex.pauljianliao.com/ppformula-model",
    [switch]$DownloadMissingModels
)

$ErrorActionPreference = "Stop"
Set-StrictMode -Version Latest

$ProjectRoot = (Resolve-Path (Join-Path $PSScriptRoot "..")).Path
$RuntimePython = Join-Path $env:APPDATA "com.visualtex.studio\ocr-runtime\python\python.exe"
$BuildCache = Join-Path $env:LOCALAPPDATA "VisualTeXBuildCache\ocr-models\windows-x64"
$BuildModelRoot = Join-Path $BuildCache "official_models"
$ArtifactRoot = Join-Path $ProjectRoot "artifacts\ocr-models\windows-x64"
$CatalogRoot = Join-Path $ProjectRoot "src-tauri\resources\ocr-models\windows-x64"
$CatalogPath = Join-Path $CatalogRoot "catalog.json"

$Models = [ordered]@{
    "PP-FormulaNet_plus-S" = [ordered]@{
        "inference.json" = "01238434e33df83588e2627f350559b576e34551d2b2ffea148345032de56c00"
        "inference.pdiparams" = "e464f94412feaa98f8791eacc84684f887b3569e30e80c52b8112e9cf7d4069b"
        "inference.yml" = "96062655d94c21d39274328dbc82c1a487e66addb8425f5a7fd5b7dfb2421ec3"
    }
    "PP-FormulaNet_plus-M" = [ordered]@{
        "inference.json" = "8333a7f650766a748e273c550d278601dd19dfeee1c4b01038ff632f134d9884"
        "inference.pdiparams" = "f16ef9b5c8227da70d3ec969a5195f4d62c1154427b883f4d6cff07633654041"
        "inference.yml" = "87b5f3d7f2b2fe553627d77b37f496608ca150ebd0ef62d362591edca47b5538"
    }
    "PP-FormulaNet_plus-L" = [ordered]@{
        "inference.json" = "ad259c4b896d99aa3479336b9121112fb40ff1ababfbf8765a3428a3b86df582"
        "inference.pdiparams" = "4245c39c181d1d21e472bc85c7434df9b23f177be46552c0542bf153addbc355"
        "inference.yml" = "afc92a2737268da0499c37b0b6741da268c369fd7424667fcfeb8fa6c7b22d30"
    }
}

function Write-Utf8NoBom([string]$Path, [string]$Content) {
    $encoding = New-Object System.Text.UTF8Encoding($false)
    [System.IO.File]::WriteAllText($Path, $Content, $encoding)
}

function Get-Sha256([string]$Path) {
    return (Get-FileHash -LiteralPath $Path -Algorithm SHA256).Hash.ToLowerInvariant()
}

function Test-ModelDirectory([string]$Directory, [System.Collections.IDictionary]$Expected) {
    if (-not (Test-Path -LiteralPath $Directory -PathType Container)) {
        return $false
    }
    $actualFiles = @(Get-ChildItem -LiteralPath $Directory -File | Select-Object -ExpandProperty Name | Sort-Object)
    $requiredFiles = @($Expected.Keys | Sort-Object)
    foreach ($name in $requiredFiles) {
        $path = Join-Path $Directory $name
        if (-not (Test-Path -LiteralPath $path -PathType Leaf)) {
            return $false
        }
        if ((Get-Sha256 $path) -ne ([string]$Expected[$name]).ToLowerInvariant()) {
            return $false
        }
    }
    return $true
}

function Find-ModelDirectory([string]$Model, [System.Collections.IDictionary]$Expected) {
    $candidates = @(
        (Join-Path $env:APPDATA "com.visualtex.studio\ocr-runtime\cache\paddlex\official_models\$Model"),
        (Join-Path $env:USERPROFILE ".paddlex\official_models\$Model"),
        (Join-Path $BuildModelRoot $Model)
    )
    foreach ($candidate in $candidates) {
        if (Test-ModelDirectory $candidate $Expected) {
            return (Resolve-Path -LiteralPath $candidate).Path
        }
    }
    return $null
}

function Download-Model([string]$Model) {
    if (-not (Test-Path -LiteralPath $RuntimePython -PathType Leaf)) {
        throw "The verified private VisualTeX OCR Python is required to prepare model assets: $RuntimePython"
    }
    New-Item -ItemType Directory -Path $BuildCache -Force | Out-Null
    $script = Join-Path $BuildCache "download-model.py"
    @'
import os
import sys
from paddleocr import FormulaRecognition

model_name = sys.argv[1]
print(f"Preparing {model_name} in explicit release-build mode", flush=True)
FormulaRecognition(model_name=model_name, device="cpu")
print(f"Prepared {model_name}", flush=True)
'@ | Set-Content -LiteralPath $script -Encoding UTF8

    $saved = @{}
    $environment = @{
        PYTHONNOUSERSITE = "1"
        PYTHONSAFEPATH = "1"
        PYTHONUTF8 = "1"
        PYTHONIOENCODING = "utf-8"
        PADDLE_PDX_CACHE_HOME = $BuildCache
        PADDLE_PDX_MODEL_SOURCE = "BOS"
        PADDLE_PDX_DISABLE_MODEL_SOURCE_CHECK = "True"
    }
    try {
        foreach ($entry in $environment.GetEnumerator()) {
            $saved[$entry.Key] = [Environment]::GetEnvironmentVariable($entry.Key, "Process")
            [Environment]::SetEnvironmentVariable($entry.Key, [string]$entry.Value, "Process")
        }
        & $RuntimePython -I -X utf8 $script $Model
        if ($LASTEXITCODE -ne 0) {
            throw "Explicit release-build download failed for $Model with exit code $LASTEXITCODE"
        }
    } finally {
        foreach ($entry in $saved.GetEnumerator()) {
            [Environment]::SetEnvironmentVariable($entry.Key, $entry.Value, "Process")
        }
    }
}

function Get-TarExecutable() {
    $windowsRoot = if ($env:SystemRoot) { $env:SystemRoot } else { "C:\Windows" }
    foreach ($directory in @("Sysnative", "System32")) {
        $candidate = Join-Path $windowsRoot "$directory\tar.exe"
        if (Test-Path -LiteralPath $candidate -PathType Leaf) {
            return $candidate
        }
    }
    throw "Windows tar.exe was not found under SystemRoot"
}

if (-not $BaseUrl.StartsWith("https://", [StringComparison]::OrdinalIgnoreCase)) {
    throw "Model catalog BaseUrl must use HTTPS"
}
$BaseUrl = $BaseUrl.TrimEnd('/')
$tar = Get-TarExecutable
New-Item -ItemType Directory -Path $ArtifactRoot -Force | Out-Null
New-Item -ItemType Directory -Path $CatalogRoot -Force | Out-Null
New-Item -ItemType Directory -Path $BuildCache -Force | Out-Null

$catalogEntries = @()
foreach ($modelEntry in $Models.GetEnumerator()) {
    $model = [string]$modelEntry.Key
    $expected = [System.Collections.IDictionary]$modelEntry.Value
    $source = Find-ModelDirectory $model $expected
    if (-not $source -and $DownloadMissingModels) {
        Download-Model $model
        $source = Find-ModelDirectory $model $expected
    }
    if (-not $source) {
        throw "Verified source files for $model are unavailable. Re-run with -DownloadMissingModels to prepare the release asset explicitly."
    }

    $workRoot = Join-Path $env:TEMP ("visualtex-model-pack-" + [guid]::NewGuid().ToString("N"))
    try {
        $packRoot = Join-Path $workRoot "visualtex-model-pack"
        $modelRoot = Join-Path $packRoot "paddlex\official_models\$model"
        New-Item -ItemType Directory -Path $modelRoot -Force | Out-Null
        $records = [ordered]@{}
        foreach ($name in $expected.Keys) {
            $sourceFile = Join-Path $source $name
            $targetFile = Join-Path $modelRoot $name
            Copy-Item -LiteralPath $sourceFile -Destination $targetFile -Force
            $item = Get-Item -LiteralPath $targetFile
            $actualHash = Get-Sha256 $targetFile
            $expectedHash = ([string]$expected[$name]).ToLowerInvariant()
            if ($actualHash -ne $expectedHash) {
                throw "SHA-256 mismatch while staging $model/$name"
            }
            $records[$name] = [ordered]@{
                name = $name
                size = $item.Length
                sha256 = $actualHash
            }
        }
        $manifest = [ordered]@{
            schemaVersion = 1
            platform = "windows"
            architecture = "x64"
            model = $model
            files = $records
        }
        Write-Utf8NoBom (Join-Path $packRoot "pack-manifest.json") ($manifest | ConvertTo-Json -Depth 8)

        $fileName = "VisualTeX_${model}_windows-x64.vtxocrmodel"
        $packagePath = Join-Path $ArtifactRoot $fileName
        $temporaryPackage = "$packagePath.tmp"
        Remove-Item -LiteralPath $temporaryPackage -Force -ErrorAction SilentlyContinue
        & $tar -czf $temporaryPackage -C $workRoot "visualtex-model-pack"
        if ($LASTEXITCODE -ne 0) {
            throw "tar.exe failed to create $fileName"
        }
        Move-Item -LiteralPath $temporaryPackage -Destination $packagePath -Force
        $package = Get-Item -LiteralPath $packagePath
        $catalogEntries += [ordered]@{
            model = $model
            url = "$BaseUrl/$fileName"
            size = $package.Length
            sha256 = Get-Sha256 $packagePath
        }
        Write-Host "Prepared ${model}: $($package.Length) bytes"
    } finally {
        Remove-Item -LiteralPath $workRoot -Recurse -Force -ErrorAction SilentlyContinue
    }
}

$catalog = [ordered]@{
    schemaVersion = 1
    platform = "windows"
    architecture = "x64"
    entries = $catalogEntries
}
Write-Utf8NoBom $CatalogPath ($catalog | ConvertTo-Json -Depth 8)
Write-Host "Prepared OCR model catalog: $CatalogPath"
Write-Host "Model packages remain outside the installer under: $ArtifactRoot"
