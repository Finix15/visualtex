param(
    [string]$OutputRoot = "",
    [switch]$Force
)
$ErrorActionPreference = "Stop"
$root = Split-Path -Parent $PSScriptRoot
$source = Join-Path $root "src-windows\VisualTeX.MathTypeSidecar"
$baseArchive = Join-Path $root "src-tauri\resources\ocr-python\windows-x64\visualtex-python-3.12.10-windows-x64.zip"
if (-not $OutputRoot) { $OutputRoot = Join-Path $root "src-windows\artifacts\mathtype-runtime\windows-x64" }
if (-not (Test-Path -LiteralPath $baseArchive)) { throw "Private Python base archive is missing: $baseArchive" }
if ((Test-Path -LiteralPath $OutputRoot) -and $Force) { Remove-Item -LiteralPath $OutputRoot -Recurse -Force }
if (-not (Test-Path -LiteralPath (Join-Path $OutputRoot "runtime-manifest.json"))) {
    New-Item -ItemType Directory -Force -Path $OutputRoot | Out-Null
    Expand-Archive -LiteralPath $baseArchive -DestinationPath $OutputRoot -Force
    $sitePackages = Join-Path $OutputRoot "Lib\site-packages"
    New-Item -ItemType Directory -Force -Path $sitePackages | Out-Null
    foreach ($wheel in Get-ChildItem -LiteralPath (Join-Path $source "wheelhouse") -Filter "*.whl") {
        $temporaryZip = Join-Path $env:TEMP ($wheel.BaseName + ".zip")
        Copy-Item -LiteralPath $wheel.FullName -Destination $temporaryZip -Force
        try { Expand-Archive -LiteralPath $temporaryZip -DestinationPath $sitePackages -Force }
        finally { Remove-Item -LiteralPath $temporaryZip -Force -ErrorAction SilentlyContinue }
    }
    Copy-Item -LiteralPath (Join-Path $source "vendor\mathtypejx") -Destination $sitePackages -Recurse -Force
    Copy-Item -LiteralPath (Join-Path $source "worker.py") -Destination (Join-Path $OutputRoot "worker.py") -Force
    Copy-Item -LiteralPath (Join-Path $source "PROVENANCE.md") -Destination (Join-Path $OutputRoot "PROVENANCE.md") -Force
    Copy-Item -LiteralPath (Join-Path $source "vendor\MATHTYPEJX-LICENSE.txt") -Destination $OutputRoot -Force
    Copy-Item -LiteralPath (Join-Path $source "vendor\MATHTYPEJX-NOTICE.txt") -Destination $OutputRoot -Force
    Copy-Item -LiteralPath (Join-Path $source "vendor\FONTMAPS-LICENSE.txt") -Destination $OutputRoot -Force
    $manifest = [ordered]@{
        schemaVersion = 1
        protocolVersion = 1
        architecture = "x64"
        pythonVersion = "3.12.10"
        mathtypejxCommit = "7d90e7274c85cf56ac28d4d15e593044693d7e70"
        dependencies = @("lxml==6.1.3", "olefile==0.47")
    }
    $manifest | ConvertTo-Json -Depth 4 | Set-Content -LiteralPath (Join-Path $OutputRoot "runtime-manifest.json") -Encoding UTF8
}
& (Join-Path $OutputRoot "python.exe") -I -X utf8 -c "import lxml, olefile, mathtypejx; print('MathType sidecar READY')"
if ($LASTEXITCODE -ne 0) { throw "Prepared MathType sidecar failed its import probe." }

function Get-StableId([string]$Prefix, [string]$Value) {
    $sha = [Security.Cryptography.SHA256]::Create()
    try { $hash = $sha.ComputeHash([Text.Encoding]::UTF8.GetBytes($Value.ToLowerInvariant())) }
    finally { $sha.Dispose() }
    return $Prefix + ([BitConverter]::ToString($hash, 0, 12).Replace('-', ''))
}
function Escape-Xml([string]$Value) { return [Security.SecurityElement]::Escape($Value) }

$runtimeRoot = [IO.Path]::GetFullPath($OutputRoot).TrimEnd('\')
$allFiles = @(Get-ChildItem -LiteralPath $runtimeRoot -Recurse -File | Sort-Object FullName)
$relativeDirectories = @($allFiles | ForEach-Object {
    $relative = $_.FullName.Substring($runtimeRoot.Length + 1)
    $directory = Split-Path -Parent $relative
    while ($directory) {
        $directory
        $directory = Split-Path -Parent $directory
    }
} | Sort-Object -Unique)
$directoryIds = @{}
foreach ($directory in $relativeDirectories) { $directoryIds[$directory] = Get-StableId "Mtd" $directory }
$children = @{}
foreach ($directory in $relativeDirectories) {
    $parent = Split-Path -Parent $directory
    if (-not $children.ContainsKey($parent)) { $children[$parent] = @() }
    $children[$parent] += $directory
}
function Add-DirectoryXml([Collections.Generic.List[string]]$Lines, [string]$Parent, [int]$Indent) {
    foreach ($directory in @($children[$Parent] | Sort-Object)) {
        $name = Split-Path -Leaf $directory
        $padding = ' ' * $Indent
        $Lines.Add("$padding<Directory Id=`"$($directoryIds[$directory])`" Name=`"$(Escape-Xml $name)`">")
        Add-DirectoryXml $Lines $directory ($Indent + 2)
        $Lines.Add("$padding</Directory>")
    }
}
$lines = [Collections.Generic.List[string]]::new()
$lines.Add('<?xml version="1.0" encoding="utf-8"?>')
$lines.Add('<Wix xmlns="http://wixtoolset.org/schemas/v4/wxs">')
$lines.Add('  <Fragment>')
$lines.Add('    <DirectoryRef Id="MathTypeRuntimeFolder">')
Add-DirectoryXml $lines '' 6
$lines.Add('    </DirectoryRef>')
$lines.Add('  </Fragment>')
$lines.Add('  <Fragment>')
$lines.Add('    <ComponentGroup Id="MathTypeRuntimeFiles">')
foreach ($file in $allFiles) {
    $relative = $file.FullName.Substring($runtimeRoot.Length + 1)
    $directory = Split-Path -Parent $relative
    $directoryId = if ($directory) { $directoryIds[$directory] } else { 'MathTypeRuntimeFolder' }
    $componentId = Get-StableId 'Mtc' $relative
    $sourcePath = '$(var.MathTypeRuntime)\' + $relative
    $lines.Add("      <Component Id=`"$componentId`" Directory=`"$directoryId`" Guid=`"*`" Bitness=`"`$(var.ComponentBitness)`">")
    $lines.Add("        <File Source=`"$(Escape-Xml $sourcePath)`" />")
    $lines.Add('      </Component>')
}
$lines.Add('    </ComponentGroup>')
$lines.Add('  </Fragment>')
$lines.Add('</Wix>')
$generatedWix = Join-Path $root 'src-windows\VisualTeX.WindowsOffice.Installer\MathTypeRuntime.Generated.wxs'
$lines | Set-Content -LiteralPath $generatedWix -Encoding UTF8
Write-Host "Generated WiX fragment for $($allFiles.Count) MathType runtime files: $generatedWix"
Write-Host "Prepared MathType sidecar: $OutputRoot"
