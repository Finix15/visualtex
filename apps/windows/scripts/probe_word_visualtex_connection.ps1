$ErrorActionPreference = "Stop"
$word = $addins = $addin = $null
try {
    $word = [Runtime.InteropServices.Marshal]::GetActiveObject("Word.Application")
    $addins = $word.COMAddIns
    Write-Output "COUNT=$($addins.Count)"
    for ($index = 1; $index -le $addins.Count; $index++) {
        $candidate = $addins.Item($index)
        try { Write-Output "ADDIN=$($candidate.ProgId)|$($candidate.Connect)" }
        finally { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($candidate) }
    }
    $addin = $addins.Item("VisualTeX.WordVsto")
    Write-Output "VISUALTEX_CONNECT=$($addin.Connect)"
} finally {
    foreach ($value in @($addin, $addins, $word)) {
        if ($null -ne $value) { try { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($value) } catch {} }
    }
}
