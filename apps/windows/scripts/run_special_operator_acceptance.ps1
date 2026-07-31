[CmdletBinding()]
param(
    [string]$AcceptanceExePath,
    [string]$ArtifactRoot
)

$ErrorActionPreference = "Stop"
Set-StrictMode -Version Latest

$root = Split-Path -Parent $PSScriptRoot
if ([string]::IsNullOrWhiteSpace($AcceptanceExePath)) {
    $AcceptanceExePath = Join-Path $root "src-windows\VisualTeX.NativeOfficeOleAcceptance\bin\x64\Release\net48\VisualTeX.NativeOfficeOleAcceptance.exe"
}
if ([string]::IsNullOrWhiteSpace($ArtifactRoot)) {
    $ArtifactRoot = Join-Path $root "build-logs\special-operator-acceptance"
}

$fixtures = @(
    Join-Path $PSScriptRoot "fixtures\windows-office-special-operators.md"
    Join-Path $PSScriptRoot "fixtures\windows-office-special-operators.tex"
    Join-Path $PSScriptRoot "fixtures\windows-office-latex-compatibility.md"
)

foreach ($path in @($AcceptanceExePath) + $fixtures) {
    if (-not (Test-Path -LiteralPath $path)) {
        throw "Special-operator acceptance input is missing: $path"
    }
}

New-Item -ItemType Directory -Force -Path $ArtifactRoot | Out-Null

Push-Location $root
try {
    & npm.cmd run test:svg-export
    if ($LASTEXITCODE -ne 0) {
        throw "SVG/MathML special-operator acceptance failed with exit code $LASTEXITCODE."
    }

    & dotnet test "src-windows\VisualTeX.WindowsOffice.Tests\VisualTeX.WindowsOffice.Tests.csproj" `
        --no-restore `
        --filter "FullyQualifiedName~WordBulkImportParserTests|FullyQualifiedName~WordOmmlTests"
    if ($LASTEXITCODE -ne 0) {
        throw "Parser/OMML special-operator acceptance failed with exit code $LASTEXITCODE."
    }

    foreach ($fixture in $fixtures) {
        $format = [IO.Path]::GetExtension($fixture).TrimStart('.').ToLowerInvariant()
        foreach ($mode in @("ole", "omml")) {
            $caseRoot = Join-Path $ArtifactRoot "$format-$mode"
            New-Item -ItemType Directory -Force -Path $caseRoot | Out-Null
            & $AcceptanceExePath --installed-word-bulk-import $fixture $mode $caseRoot
            if ($LASTEXITCODE -ne 0) {
                throw "Real Word special-operator acceptance failed: format=$format mode=$mode exit=$LASTEXITCODE."
            }
        }
    }
}
finally {
    Pop-Location
}

Write-Host "Special-operator acceptance passed: MathML/SVG + parser + OMML + Markdown/LaTeX x OLE/OMML real Word matrix."
Write-Host "Artifacts: $ArtifactRoot"
