[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [ValidateSet("Word", "PowerPoint")]
    [string]$HostName,
    [Parameter(Mandatory = $true)]
    [string]$ProgId,
    [Parameter(Mandatory = $true)]
    [string]$ResultPath,
    [int]$ConnectTimeoutSeconds = 10
)

$ErrorActionPreference = "Stop"
$application = $null
$document = $null
$presentation = $null
$addIns = $null
$addIn = $null
$inventory = @()
$connectAttempted = $false
$stage = "create-hidden-application"

function Release-ComObject([object]$Value) {
    if ($null -ne $Value -and [Runtime.InteropServices.Marshal]::IsComObject($Value)) {
        try { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($Value) } catch { }
    }
}

function Write-ProbeResult([object]$Result) {
    $Result | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath $ResultPath -Encoding UTF8
}

try {
    if ($HostName -eq "Word") {
        $application = New-Object -ComObject Word.Application
        $application.Visible = $false
        $application.DisplayAlerts = 0
        $stage = "create-temporary-document"
        $document = $application.Documents.Add()
    } else {
        $application = New-Object -ComObject PowerPoint.Application
        $stage = "create-temporary-presentation"
        $presentation = $application.Presentations.Add(0)
    }

    $stage = "enumerate-com-addins"
    $deadline = [DateTimeOffset]::UtcNow.AddSeconds([Math]::Max(1, $ConnectTimeoutSeconds))
    do {
        Release-ComObject $addIn
        $addIn = $null
        Release-ComObject $addIns
        $addIns = $application.COMAddIns
        $inventory = @()
        $count = 0
        try { $count = [int]$addIns.Count } catch { }
        for ($index = 1; $index -le $count; $index++) {
            $candidate = $null
            try {
                $candidate = $addIns.Item($index)
                $candidateProgId = [string]$candidate.ProgId
                $candidateConnected = $false
                try { $candidateConnected = [bool]$candidate.Connect } catch { }
                $inventory += [pscustomobject]@{
                    index = $index
                    progId = $candidateProgId
                    description = [string]$candidate.Description
                    connected = $candidateConnected
                }
                if ([string]::Equals($candidateProgId, $ProgId, [StringComparison]::OrdinalIgnoreCase)) {
                    $addIn = $candidate
                    $candidate = $null
                }
            } finally {
                Release-ComObject $candidate
            }
        }
        if ($null -ne $addIn) { break }
        Start-Sleep -Milliseconds 250
    } while ([DateTimeOffset]::UtcNow -lt $deadline)

    if ($null -eq $addIn) {
        throw "$HostName did not enumerate the VisualTeX add-in."
    }

    $stage = "read-connect"
    $connected = [bool]$addIn.Connect
    if (-not $connected) {
        $stage = "connect-addin"
        $connectAttempted = $true
        $addIn.Connect = $true
        $deadline = [DateTimeOffset]::UtcNow.AddSeconds(5)
        do {
            Start-Sleep -Milliseconds 200
            $connected = [bool]$addIn.Connect
        } while (-not $connected -and [DateTimeOffset]::UtcNow -lt $deadline)
    }

    Write-ProbeResult ([pscustomobject]@{
        host = $HostName
        progId = $ProgId
        enumerated = $true
        connected = $connected
        connectAttempted = $connectAttempted
        description = [string]$addIn.Description
        discoveredProgIds = @($inventory | ForEach-Object { $_.progId })
        inventory = $inventory
        stage = "complete"
        startupMode = "hidden-com-worker"
        error = if ($connected) { "" } else { "$HostName enumerated the VisualTeX add-in, but COMAddIn.Connect remained false." }
    })
    exit $(if ($connected) { 0 } else { 1 })
} catch {
    Write-ProbeResult ([pscustomobject]@{
        host = $HostName
        progId = $ProgId
        enumerated = $null -ne $addIn
        connected = $false
        connectAttempted = $connectAttempted
        description = if ($null -ne $addIn) { [string]$addIn.Description } else { "" }
        discoveredProgIds = @($inventory | ForEach-Object { $_.progId })
        inventory = $inventory
        stage = $stage
        startupMode = "hidden-com-worker"
        error = $_.Exception.Message
    })
    exit 1
} finally {
    if ($null -ne $document) { try { $document.Close(0) } catch { } }
    if ($null -ne $presentation) { try { $presentation.Close() } catch { } }
    if ($null -ne $application) { try { $application.Quit() } catch { } }
    Release-ComObject $addIn
    Release-ComObject $addIns
    Release-ComObject $document
    Release-ComObject $presentation
    Release-ComObject $application
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
