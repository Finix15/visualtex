[CmdletBinding()]
param(
    [ValidateSet("auto", "x86", "x64")]
    [string]$OfficePlatform = "auto",
    [string]$VisualTeXPath,
    [switch]$CompanionOnly,
    [string]$ReportPath
)

$ErrorActionPreference = "Stop"
Add-Type -AssemblyName System.Net.Http
$logRoot = Join-Path $env:LOCALAPPDATA "VisualTeX\office\install-logs"
New-Item -Path $logRoot -ItemType Directory -Force | Out-Null
$stamp = Get-Date -Format "yyyyMMdd-HHmmss"
if ([string]::IsNullOrWhiteSpace($ReportPath)) {
    $ReportPath = Join-Path $logRoot "office-runtime-$stamp.json"
}
$script:checks = New-Object System.Collections.Generic.List[object]
$script:failures = New-Object System.Collections.Generic.List[string]
$script:resolvedOfficePlatform = $null
$officeMsiUpgradeCode = "{A81B4BF7-0E51-45CE-A5AA-5E28F6944F42}"
$officeMsiDisplayName = "VisualTeX Windows Office Integration"

function Add-Check {
    param([string]$Name, [bool]$Passed, [string]$Details)
    [void]$script:checks.Add([pscustomobject]@{
        name = $Name
        passed = $Passed
        details = $Details
    })
    if ($Passed) { Write-Host "[PASS] $Name - $Details" }
    else {
        Write-Warning "[FAIL] $Name - $Details"
        [void]$script:failures.Add("${Name}: $Details")
    }
}

function Get-RegistryView([string]$Architecture) {
    if ($Architecture -eq "x86") { return [Microsoft.Win32.RegistryView]::Registry32 }
    return [Microsoft.Win32.RegistryView]::Registry64
}

function Get-RegistryValue {
    param(
        [Microsoft.Win32.RegistryHive]$Hive,
        [string]$SubKey,
        [string]$Name,
        [string]$Architecture
    )
    $baseKey = [Microsoft.Win32.RegistryKey]::OpenBaseKey($Hive, (Get-RegistryView $Architecture))
    try {
        $key = $baseKey.OpenSubKey($SubKey, $false)
        if ($null -eq $key) { return $null }
        try {
            $valueName = if ($Name -eq "(default)") { "" } else { $Name }
            if ($valueName -notin @($key.GetValueNames())) { return $null }
            return $key.GetValue(
                $valueName,
                $null,
                [Microsoft.Win32.RegistryValueOptions]::DoNotExpandEnvironmentNames)
        } finally {
            $key.Dispose()
        }
    } finally {
        $baseKey.Dispose()
    }
}

function Resolve-OfficePlatform([string]$Requested) {
    $configuration = Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration" -ErrorAction SilentlyContinue
    $detected = if ($configuration.Platform -in @("x86", "x64")) { [string]$configuration.Platform } else { $null }
    if ($Requested -ne "auto") {
        if ($detected -and $detected -ne $Requested) {
            throw "Requested Office platform '$Requested' does not match installed Click-to-Run Office '$detected'."
        }
        return $Requested
    }
    if ($detected) { return $detected }
    foreach ($architecture in @("x64", "x86")) {
        $word = Get-RegistryValue ([Microsoft.Win32.RegistryHive]::LocalMachine) "SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\WINWORD.EXE" "(default)" $architecture
        $powerPoint = Get-RegistryValue ([Microsoft.Win32.RegistryHive]::LocalMachine) "SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\POWERPNT.EXE" "(default)" $architecture
        if ($word -and $powerPoint) { return $architecture }
    }
    throw "Unable to determine the installed Office architecture."
}

function Test-CertificateTrusted {
    $modeKey = "HKCU:\Software\VisualTeX\OfficeIntegration"
    $thumbprint = (Get-ItemProperty -LiteralPath $modeKey -Name CertificateThumbprint -ErrorAction SilentlyContinue).CertificateThumbprint
    $certificatePath = (Get-ItemProperty -LiteralPath $modeKey -Name CertificatePath -ErrorAction SilentlyContinue).CertificatePath
    if ([string]::IsNullOrWhiteSpace([string]$thumbprint) -or
        [string]::IsNullOrWhiteSpace([string]$certificatePath) -or
        -not (Test-Path -LiteralPath ([string]$certificatePath) -PathType Leaf)) {
        return [pscustomobject]@{ passed = $false; details = "Certificate registry values or certificate file are missing." }
    }
    try {
        $certificate = [Security.Cryptography.X509Certificates.X509Certificate2]::new([string]$certificatePath)
        $store = [Security.Cryptography.X509Certificates.X509Store]::new(
            [Security.Cryptography.X509Certificates.StoreName]::Root,
            [Security.Cryptography.X509Certificates.StoreLocation]::CurrentUser)
        $store.Open([Security.Cryptography.X509Certificates.OpenFlags]::ReadOnly)
        try {
            $trusted = @($store.Certificates | Where-Object { $_.Thumbprint -eq $certificate.Thumbprint }).Count -gt 0
        } finally {
            $store.Close()
        }
        return [pscustomobject]@{
            passed = $trusted -and $certificate.Thumbprint -eq ([string]$thumbprint).Replace(" ", "")
            details = "$($certificate.Thumbprint); $certificatePath"
        }
    } catch {
        return [pscustomobject]@{ passed = $false; details = $_.Exception.Message }
    }
}

function Get-VstoRuntimeState([string]$Architecture) {
    $subKey = "SOFTWARE\Microsoft\VSTO Runtime Setup\v4R"
    foreach ($registryArchitecture in @($Architecture, "x86", "x64") | Select-Object -Unique) {
        $install = Get-RegistryValue ([Microsoft.Win32.RegistryHive]::LocalMachine) $subKey "Install" $registryArchitecture
        $clr40 = Get-RegistryValue ([Microsoft.Win32.RegistryHive]::LocalMachine) $subKey "VSTORFeature_CLR40" $registryArchitecture
        $version = Get-RegistryValue ([Microsoft.Win32.RegistryHive]::LocalMachine) $subKey "Version" $registryArchitecture
        $installed =
            ($null -ne $install -and [int]$install -eq 1) -or
            ($null -ne $clr40 -and [int]$clr40 -eq 1) -or
            (-not [string]::IsNullOrWhiteSpace([string]$version))
        if ($installed) {
            return [pscustomobject]@{
                installed = $true
                install = $install
                clr40 = $clr40
                version = $version
                registryView = $registryArchitecture
                key = "HKLM\$subKey"
            }
        }
    }
    return [pscustomobject]@{
        installed = $false
        install = $null
        clr40 = $null
        version = $null
        registryView = "none"
        key = "HKLM\$subKey"
    }
}

function Get-MsiInstalledStateOnce {
    $codes = New-Object System.Collections.Generic.HashSet[string] ([StringComparer]::OrdinalIgnoreCase)
    $sources = New-Object System.Collections.Generic.List[string]

    $installer = $null
    try {
        $installer = New-Object -ComObject WindowsInstaller.Installer
        $related = $installer.GetType().InvokeMember(
            "RelatedProducts",
            [Reflection.BindingFlags]::GetProperty,
            $null,
            $installer,
            @($officeMsiUpgradeCode))
        foreach ($code in @($related)) {
            if (-not [string]::IsNullOrWhiteSpace([string]$code)) {
                [void]$codes.Add([string]$code)
                if (-not $sources.Contains("WindowsInstaller.RelatedProducts")) {
                    [void]$sources.Add("WindowsInstaller.RelatedProducts")
                }
            }
        }
    } catch {
        # The uninstall registry fallback below is intentionally independent of COM.
    } finally {
        if ($null -ne $installer) {
            [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($installer)
        }
    }

    foreach ($root in @(
        "HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall",
        "HKCU:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall",
        "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall",
        "HKLM:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall"
    )) {
        if (-not (Test-Path -LiteralPath $root)) { continue }
        foreach ($key in Get-ChildItem -LiteralPath $root -ErrorAction SilentlyContinue) {
            $item = Get-ItemProperty -LiteralPath $key.PSPath -ErrorAction SilentlyContinue
            if ([string]$item.DisplayName -eq $officeMsiDisplayName) {
                [void]$codes.Add([string]$key.PSChildName)
                if (-not $sources.Contains($root)) { [void]$sources.Add($root) }
            }
        }
    }

    $productCodes = @($codes)
    return [pscustomobject]@{
        installed = $productCodes.Count -gt 0
        productCode = if ($productCodes.Count -gt 0) { [string]$productCodes[0] } else { "" }
        productCodes = $productCodes
        source = $sources -join ", "
    }
}

function Test-MsiInstalled([int]$TimeoutSeconds = 10) {
    $deadline = [DateTimeOffset]::UtcNow.AddSeconds($TimeoutSeconds)
    $lastState = Get-MsiInstalledStateOnce
    while (-not $lastState.installed -and [DateTimeOffset]::UtcNow -lt $deadline) {
        Start-Sleep -Milliseconds 250
        $lastState = Get-MsiInstalledStateOnce
    }
    return $lastState
}

function Get-LoadBehavior([ValidateSet("Word", "PowerPoint")][string]$HostName, [string]$Architecture) {
    $progId = if ($HostName -eq "Word") { "VisualTeX.WordVsto" } else { "VisualTeX.PowerPointVsto" }
    return Get-RegistryValue ([Microsoft.Win32.RegistryHive]::CurrentUser) "Software\Microsoft\Office\$HostName\Addins\$progId" "LoadBehavior" $Architecture
}

function Get-DisabledItems([ValidateSet("Word", "PowerPoint")][string]$HostName) {
    $results = @()
    foreach ($bucket in @("DisabledItems", "StartupItems")) {
        $key = "HKCU:\Software\Microsoft\Office\16.0\$HostName\Resiliency\$bucket"
        if (-not (Test-Path -LiteralPath $key)) { continue }
        $item = Get-ItemProperty -LiteralPath $key -ErrorAction SilentlyContinue
        foreach ($property in $item.PSObject.Properties | Where-Object { $_.Name -notmatch '^PS' }) {
            if ($property.Value -is [byte[]]) {
                $unicode = [Text.Encoding]::Unicode.GetString([byte[]]$property.Value).Trim([char]0)
                $ascii = [Text.Encoding]::ASCII.GetString([byte[]]$property.Value).Trim([char]0)
                $results += [pscustomobject]@{
                    bucket = $bucket
                    name = $property.Name
                    unicode = $unicode
                    ascii = $ascii
                    mentionsVisualTeX = ($unicode -match 'VisualTeX' -or $ascii -match 'VisualTeX')
                }
            }
        }
    }
    return @($results)
}

function Get-RecentOfficeLoadEvents {
    $start = (Get-Date).AddMinutes(-15)
    try {
        return @(Get-WinEvent -FilterHashtable @{ LogName = "Application"; StartTime = $start } -MaxEvents 150 -ErrorAction Stop |
            Where-Object {
                $_.ProviderName -match 'VSTO|\.NET Runtime|Office|Application Error' -or
                $_.Message -match 'VisualTeX|VisualTeX\.WordVsto|VisualTeX\.PowerPointVsto'
            } |
            Select-Object -First 30 TimeCreated, ProviderName, Id, LevelDisplayName, Message)
    } catch {
        return @([pscustomobject]@{
            TimeCreated = Get-Date
            ProviderName = "Get-WinEvent"
            Id = 0
            LevelDisplayName = "Warning"
            Message = $_.Exception.Message
        })
    }
}

function Test-TcpPort {
    param(
        [string]$HostName,
        [int]$Port,
        [int]$TimeoutMilliseconds = 2000
    )
    $client = New-Object System.Net.Sockets.TcpClient
    try {
        $async = $client.BeginConnect($HostName, $Port, $null, $null)
        if (-not $async.AsyncWaitHandle.WaitOne($TimeoutMilliseconds, $false)) {
            return $false
        }
        $client.EndConnect($async)
        return $true
    } catch {
        return $false
    } finally {
        $client.Close()
    }
}

function Get-RemoteCertificateThumbprint {
    param([string]$HostName, [int]$Port)
    $client = New-Object System.Net.Sockets.TcpClient
    $stream = $null
    try {
        $client.Connect($HostName, $Port)
        $callback = [Net.Security.RemoteCertificateValidationCallback]{
            param($sender, $certificate, $chain, $sslPolicyErrors)
            return $true
        }
        $stream = New-Object System.Net.Security.SslStream($client.GetStream(), $false, $callback)
        $stream.AuthenticateAsClient("localhost")
        $remoteCertificate = [Security.Cryptography.X509Certificates.X509Certificate2]::new($stream.RemoteCertificate)
        return $remoteCertificate.Thumbprint
    } finally {
        if ($null -ne $stream) { $stream.Dispose() }
        $client.Close()
    }
}

function Get-ExceptionChain([Exception]$Exception) {
    $items = @()
    for ($current = $Exception; $null -ne $current; $current = $current.InnerException) {
        $items += [pscustomobject]@{
            type = $current.GetType().FullName
            message = $current.Message
            hresult = ('0x{0:X8}' -f $current.HResult)
            stackTrace = $current.StackTrace
        }
    }
    return @($items)
}

function Get-ProcessIdentity {
    param(
        [int]$ProcessId,
        [int]$PathRetryMilliseconds = 3000
    )
    $deadline = [DateTimeOffset]::UtcNow.AddMilliseconds($PathRetryMilliseconds)
    $processName = ""
    $processPath = ""
    do {
        $process = Get-Process -Id $ProcessId -ErrorAction SilentlyContinue
        if ($null -eq $process) { break }
        $processName = [string]$process.ProcessName
        try { $processPath = [string]$process.Path } catch { $processPath = "" }
        if ([string]::IsNullOrWhiteSpace($processPath)) {
            try {
                $cimProcess = Get-CimInstance Win32_Process -Filter "ProcessId = $ProcessId" -ErrorAction SilentlyContinue
                $processPath = [string]$cimProcess.ExecutablePath
            } catch { $processPath = "" }
        }
        if (-not [string]::IsNullOrWhiteSpace($processPath)) { break }
        Start-Sleep -Milliseconds 100
    } while ([DateTimeOffset]::UtcNow -lt $deadline)
    return [pscustomobject]@{
        processId = $ProcessId
        processName = $processName
        processPath = $processPath
    }
}

function Get-PortOwner([int]$Port, [int]$PathRetryMilliseconds = 3000) {
    $processId = $null
    $source = "none"
    try {
        $connection = Get-NetTCPConnection -State Listen -LocalPort $Port -ErrorAction Stop |
            Select-Object -First 1
        if ($null -ne $connection) {
            $processId = [int]$connection.OwningProcess
            $source = "Get-NetTCPConnection"
        }
    } catch { }

    if ($null -eq $processId) {
        try {
            foreach ($line in (& netstat.exe -ano -p tcp 2>$null)) {
                if ($line -notmatch 'LISTENING') { continue }
                $columns = @($line.Trim() -split '\s+')
                if ($columns.Count -lt 5 -or $columns[1] -notmatch ":$Port$") { continue }
                $processId = [int]$columns[-1]
                $source = "netstat"
                break
            }
        } catch { }
    }

    if ($null -ne $processId) {
        $identity = Get-ProcessIdentity $processId $PathRetryMilliseconds
        return [pscustomobject]@{
            processId = $identity.processId
            processName = $identity.processName
            processPath = $identity.processPath
            source = $source
        }
    }
    return [pscustomobject]@{
        processId = $null
        processName = ""
        processPath = ""
        source = "none"
    }
}

function Test-CompanionRuntime {
    param(
        [object]$IntegrationState,
        [string]$ExplicitVisualTeXPath
    )

    $errors = New-Object System.Collections.Generic.List[string]
    $stage = "configuration"
    $registeredExecutable = [string]$IntegrationState.ExecutablePath
    $executable = if ([string]::IsNullOrWhiteSpace($ExplicitVisualTeXPath)) {
        $registeredExecutable
    } else {
        $ExplicitVisualTeXPath.Trim().Trim('"')
    }
    if (-not [string]::IsNullOrWhiteSpace($executable) -and (Test-Path -LiteralPath $executable -PathType Leaf)) {
        $executable = (Resolve-Path -LiteralPath $executable).Path
    }
    if (-not [string]::IsNullOrWhiteSpace($registeredExecutable) -and (Test-Path -LiteralPath $registeredExecutable -PathType Leaf)) {
        $registeredExecutable = (Resolve-Path -LiteralPath $registeredExecutable).Path
    }
    $appDataRoot = [string]$IntegrationState.AppDataRoot
    $certificatePath = [string]$IntegrationState.CertificatePath
    $expectedThumbprint = ([string]$IntegrationState.CertificateThumbprint).Replace(" ", "")
    $port = [int]$IntegrationState.CompanionPort
    $expectedProtocol = [int]$IntegrationState.ProtocolVersion
    $installJsonPath = Join-Path $appDataRoot "office\install.json"
    $companionLogRoot = Join-Path $env:LOCALAPPDATA "VisualTeX\office\logs"
    $startupLog = Join-Path $companionLogRoot "startup.log"
    $companionLog = Join-Path $companionLogRoot "companion.log"
    $startedProcessId = $null
    $startedProcessExitCode = $null
    $processRunning = $false
    $matchingProcess = $null
    $portListening = $false
    $portOwner = [pscustomobject]@{ processId = $null; processName = ""; processPath = ""; source = "none" }
    $httpsHealthy = $false
    $protocolMatches = $false
    $certificateMatches = $false
    $localThumbprint = ""
    $remoteThumbprint = ""
    $tlsPolicyErrors = ""
    $healthRaw = ""
    $health = $null
    $exceptionChain = @()

    if ([string]::IsNullOrWhiteSpace($executable)) {
        [void]$errors.Add("ExecutablePath is missing. Supply -VisualTeXPath and repair the integration registry state.")
    } elseif (-not (Test-Path -LiteralPath $executable -PathType Leaf)) {
        [void]$errors.Add("The exact VisualTeX executable is missing: $executable")
    }
    if (-not [string]::IsNullOrWhiteSpace($ExplicitVisualTeXPath) -and
        -not [string]::IsNullOrWhiteSpace($registeredExecutable) -and
        -not [string]::Equals(
            [IO.Path]::GetFullPath($executable),
            [IO.Path]::GetFullPath($registeredExecutable),
            [StringComparison]::OrdinalIgnoreCase)) {
        [void]$errors.Add("Explicit VisualTeXPath '$executable' does not match registered ExecutablePath '$registeredExecutable'.")
    }
    if ([string]::IsNullOrWhiteSpace($appDataRoot)) {
        [void]$errors.Add("AppDataRoot is missing from HKCU\Software\VisualTeX\OfficeIntegration.")
    } elseif (-not (Test-Path -LiteralPath $appDataRoot -PathType Container)) {
        [void]$errors.Add("The registered AppDataRoot does not exist: $appDataRoot")
    }
    if ($port -le 0 -or $port -gt 65535) {
        [void]$errors.Add("CompanionPort is missing or invalid: $port")
    }
    if ($expectedProtocol -le 0) {
        [void]$errors.Add("ProtocolVersion is missing or invalid: $expectedProtocol")
    }
    if (-not (Test-Path -LiteralPath $installJsonPath -PathType Leaf)) {
        [void]$errors.Add("install.json is missing: $installJsonPath")
    } else {
        try {
            $installState = Get-Content -LiteralPath $installJsonPath -Raw | ConvertFrom-Json
            if ([string]::IsNullOrWhiteSpace([string]$installState.installToken) -or
                [string]$installState.installToken.Length -ne 64) {
                [void]$errors.Add("install.json contains an invalid installToken.")
            }
            if ([int]$installState.port -ne $port) {
                [void]$errors.Add("install.json port=$($installState.port) does not match registry CompanionPort=$port.")
            }
            if ([int]$installState.protocolVersion -ne $expectedProtocol) {
                [void]$errors.Add("install.json protocolVersion=$($installState.protocolVersion) does not match registry ProtocolVersion=$expectedProtocol.")
            }
        } catch {
            $exceptionChain = @(Get-ExceptionChain $_.Exception)
            [void]$errors.Add("install.json is invalid: $($_.Exception.Message)")
        }
    }

    if ($errors.Count -eq 0) {
        # Inspect the fixed companion port before starting VisualTeX. This
        # prevents a foreign listener from being misreported as a generic
        # background-process timeout or early-exit failure.
        $portListening = Test-TcpPort "127.0.0.1" $port 750
        if ($portListening) {
            $portOwner = Get-PortOwner $port
            if (-not [string]::IsNullOrWhiteSpace([string]$portOwner.processPath) -and
                -not [string]::Equals(
                    [IO.Path]::GetFullPath([string]$portOwner.processPath),
                    [IO.Path]::GetFullPath($executable),
                    [StringComparison]::OrdinalIgnoreCase)) {
                $stage = "port-conflict"
                [void]$errors.Add("Port $port is already occupied by '$($portOwner.processPath)' PID=$($portOwner.processId) process='$($portOwner.processName)', not by '$executable'.")
            } elseif ($null -ne $portOwner.processId) {
                $matchingProcess = Get-Process -Id ([int]$portOwner.processId) -ErrorAction SilentlyContinue
            }
        }
    }

    if ($errors.Count -eq 0 -and -not $portListening) {
        $stage = "process-start"
        foreach ($process in @(Get-Process -ErrorAction SilentlyContinue)) {
            $processPath = $null
            try { $processPath = $process.Path } catch { continue }
            if (-not [string]::IsNullOrWhiteSpace([string]$processPath) -and
                [string]::Equals(
                    [IO.Path]::GetFullPath([string]$processPath),
                    [IO.Path]::GetFullPath($executable),
                    [StringComparison]::OrdinalIgnoreCase)) {
                $matchingProcess = $process
                break
            }
        }
        if ($null -eq $matchingProcess) {
            try {
                $matchingProcess = Start-Process -FilePath $executable -ArgumentList "--office-background" -WindowStyle Hidden -PassThru
                $startedProcessId = $matchingProcess.Id
            } catch {
                $exceptionChain = @(Get-ExceptionChain $_.Exception)
                [void]$errors.Add("Unable to start the exact VisualTeX executable '$executable': $($_.Exception.Message)")
            }
        }
    }

    if ($errors.Count -eq 0 -and -not $portListening) {
        $stage = "port-listen"
        $deadline = [DateTimeOffset]::UtcNow.AddSeconds(20)
        do {
            if ($null -ne $matchingProcess -and $matchingProcess.HasExited) {
                $startedProcessExitCode = $matchingProcess.ExitCode
                [void]$errors.Add("VisualTeX background PID=$($matchingProcess.Id) exited with code $startedProcessExitCode before the companion became ready.")
                break
            }
            $portListening = Test-TcpPort "127.0.0.1" $port 750
            if ($portListening) { break }
            Start-Sleep -Milliseconds 250
        } while ([DateTimeOffset]::UtcNow -lt $deadline)
        if (-not $portListening) {
            [void]$errors.Add("No process listened on 127.0.0.1:$port within 20 seconds. Check '$startupLog' and '$companionLog'.")
        }
    }

    if ($portListening -and $errors.Count -eq 0) {
        $portOwner = Get-PortOwner $port
        if (-not [string]::IsNullOrWhiteSpace([string]$portOwner.processPath) -and
            -not [string]::Equals(
                [IO.Path]::GetFullPath([string]$portOwner.processPath),
                [IO.Path]::GetFullPath($executable),
                [StringComparison]::OrdinalIgnoreCase)) {
            $stage = "port-conflict"
            [void]$errors.Add("Port $port is owned by '$($portOwner.processPath)' PID=$($portOwner.processId), not by '$executable'.")
        } elseif ($null -eq $matchingProcess -and $null -ne $portOwner.processId) {
            $matchingProcess = Get-Process -Id ([int]$portOwner.processId) -ErrorAction SilentlyContinue
        }
    }
    if ($null -ne $matchingProcess -and -not $matchingProcess.HasExited) {
        $processRunning = $true
    }
    $actualProcessId = if ($null -ne $matchingProcess) { $matchingProcess.Id } else { $portOwner.processId }

    if ($portListening -and $errors.Count -eq 0) {
        $stage = "https-handshake"
        $handler = [System.Net.Http.HttpClientHandler]::new()
        $handler.UseProxy = $false
        $handler.Proxy = $null
        $client = [System.Net.Http.HttpClient]::new($handler)
        $client.Timeout = [TimeSpan]::FromSeconds(5)
        try {
            # Read the actual server certificate through a separate TLS probe.
            # The HttpClient request itself intentionally uses the Windows trust
            # chain so an untrusted current-user Root certificate fails normally.
            $remoteThumbprint = (Get-RemoteCertificateThumbprint "127.0.0.1" $port).Replace(" ", "")
            $response = $client.GetAsync("https://127.0.0.1:$port/health").GetAwaiter().GetResult()
            $healthRaw = $response.Content.ReadAsStringAsync().GetAwaiter().GetResult()
            $tlsPolicyErrors = "None"
            if (-not $response.IsSuccessStatusCode) {
                [void]$errors.Add("The companion /health endpoint returned HTTP $([int]$response.StatusCode): $healthRaw")
            } else {
                $health = $healthRaw | ConvertFrom-Json
                $httpsHealthy = [bool]$health.ok
                $protocolMatches = [int]$health.protocolVersion -eq $expectedProtocol
                if (-not $httpsHealthy) {
                    [void]$errors.Add("The companion HTTPS /health endpoint returned ok=false: $healthRaw")
                }
                if (-not $protocolMatches) {
                    [void]$errors.Add("Companion protocolVersion=$($health.protocolVersion) does not match expected $expectedProtocol.")
                }
            }
        } catch {
            $exceptionChain = @(Get-ExceptionChain $_.Exception)
            if ([string]::IsNullOrWhiteSpace($tlsPolicyErrors)) { $tlsPolicyErrors = "Windows trust-chain validation failed or the TLS handshake did not complete." }
            [void]$errors.Add("HTTPS /health failed with proxy disabled: $($_.Exception.Message); TLS=$tlsPolicyErrors")
        } finally {
            $client.Dispose()
            $handler.Dispose()
        }
    }

    if ($portListening -and $errors.Count -eq 0) {
        $stage = "certificate-match"
        if ([string]::IsNullOrWhiteSpace($expectedThumbprint)) {
            [void]$errors.Add("CertificateThumbprint is missing from HKCU\Software\VisualTeX\OfficeIntegration.")
        } elseif ([string]::IsNullOrWhiteSpace($certificatePath) -or
            -not (Test-Path -LiteralPath $certificatePath -PathType Leaf)) {
            [void]$errors.Add("The registered companion certificate file is missing: $certificatePath")
        } else {
            try {
                $localCertificate = [Security.Cryptography.X509Certificates.X509Certificate2]::new($certificatePath)
                $localThumbprint = $localCertificate.Thumbprint.Replace(" ", "")
                if ([string]::IsNullOrWhiteSpace($remoteThumbprint)) {
                    $remoteThumbprint = (Get-RemoteCertificateThumbprint "127.0.0.1" $port).Replace(" ", "")
                }
                $certificateMatches =
                    [string]::Equals($localThumbprint, $expectedThumbprint, [StringComparison]::OrdinalIgnoreCase) -and
                    [string]::Equals($remoteThumbprint, $expectedThumbprint, [StringComparison]::OrdinalIgnoreCase)
                if (-not $certificateMatches) {
                    [void]$errors.Add("Certificate mismatch: registry=$expectedThumbprint; file=$localThumbprint; server=$remoteThumbprint")
                }
            } catch {
                $exceptionChain = @(Get-ExceptionChain $_.Exception)
                [void]$errors.Add("Unable to inspect the companion certificate: $($_.Exception.Message)")
            }
        }
    }

    if ($protocolMatches -and $certificateMatches -and $errors.Count -eq 0) { $stage = "complete" }
    return [pscustomobject]@{
        stage = $stage
        executable = $executable
        registeredExecutable = $registeredExecutable
        appDataRoot = $appDataRoot
        installJsonPath = $installJsonPath
        startupLog = $startupLog
        companionLog = $companionLog
        port = $port
        expectedProtocol = $expectedProtocol
        processId = $actualProcessId
        startedProcessId = $startedProcessId
        startedProcessExitCode = $startedProcessExitCode
        processRunning = $processRunning
        portListening = $portListening
        portOwner = $portOwner
        httpsHealthy = $httpsHealthy
        certificateMatches = $certificateMatches
        protocolMatches = $protocolMatches
        certificatePath = $certificatePath
        expectedThumbprint = $expectedThumbprint
        localThumbprint = $localThumbprint
        remoteThumbprint = $remoteThumbprint
        tlsPolicyErrors = $tlsPolicyErrors
        healthRaw = $healthRaw
        health = $health
        exceptionChain = @($exceptionChain)
        errors = @($errors)
    }
}

function Test-OleLocalServer([string]$Architecture) {
    $clsid = "{8FF7F5AA-0D60-48D5-ADBD-65A64B4C827B}"
    $subKey = "Software\Classes\CLSID\$clsid\LocalServer32"
    $server = [string](Get-RegistryValue ([Microsoft.Win32.RegistryHive]::CurrentUser) $subKey "ServerExecutable" $Architecture)
    if ([string]::IsNullOrWhiteSpace($server)) {
        $server = [string](Get-RegistryValue ([Microsoft.Win32.RegistryHive]::CurrentUser) $subKey "(default)" $Architecture)
        $server = $server.Trim('"')
    }
    if ([string]::IsNullOrWhiteSpace($server) -or
        -not (Test-Path -LiteralPath $server -PathType Leaf)) {
        return [pscustomobject]@{
            healthy = $false
            server = $server
            clsid = $clsid
            embeddingProbe = $false
            error = "The registered OLE LocalServer executable is missing."
        }
    }

    $process = $null
    try {
        $process = Start-Process -FilePath $server -ArgumentList "-Embedding" -WindowStyle Hidden -PassThru
        Start-Sleep -Milliseconds 900
        if ($process.HasExited -and $process.ExitCode -ne 0) {
            return [pscustomobject]@{
                healthy = $false
                server = $server
                clsid = $clsid
                embeddingProbe = $false
                error = "The OLE LocalServer exited with code $($process.ExitCode) during the -Embedding probe."
            }
        }
        return [pscustomobject]@{
            healthy = $true
            server = $server
            clsid = $clsid
            embeddingProbe = $true
            error = ""
        }
    } catch {
        return [pscustomobject]@{
            healthy = $false
            server = $server
            clsid = $clsid
            embeddingProbe = $false
            error = $_.Exception.ToString()
        }
    } finally {
        if ($null -ne $process -and -not $process.HasExited) {
            Stop-Process -Id $process.Id -Force -ErrorAction SilentlyContinue
        }
    }
}

function Test-OfficeComAddIn {
    param(
        [ValidateSet("Word", "PowerPoint")][string]$HostName,
        [string]$ProgId
    )
    $application = $null
    $addIn = $null
    try {
        $comType = if ($HostName -eq "Word") { "Word.Application" } else { "PowerPoint.Application" }
        $application = New-Object -ComObject $comType
        Start-Sleep -Milliseconds 1200
        $addIn = $application.COMAddIns.Item($ProgId)
        $connected = [bool]$addIn.Connect
        return [pscustomobject]@{
            host = $HostName
            progId = $ProgId
            connected = $connected
            description = [string]$addIn.Description
            error = ""
        }
    } catch {
        return [pscustomobject]@{
            host = $HostName
            progId = $ProgId
            connected = $false
            description = ""
            error = $_.Exception.ToString()
        }
    } finally {
        if ($null -ne $application) {
            try { $application.Quit() } catch { }
        }
        if ($null -ne $addIn -and [Runtime.InteropServices.Marshal]::IsComObject($addIn)) {
            [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($addIn)
        }
        if ($null -ne $application -and [Runtime.InteropServices.Marshal]::IsComObject($application)) {
            [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($application)
        }
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
    }
}

function Write-Report {
    param(
        [object]$WordResult,
        [object]$PowerPointResult,
        [object]$CertificateState,
        [object]$VstoRuntime,
        [object]$MsiState,
        [object]$OleState,
        [object]$CompanionState,
        [object[]]$WordDisabledItems,
        [object[]]$PowerPointDisabledItems,
        [object[]]$Events
    )
    $report = [ordered]@{
        schemaVersion = 1
        generatedAt = [DateTimeOffset]::Now.ToString("o")
        officePlatform = $script:resolvedOfficePlatform
        succeeded = $script:failures.Count -eq 0
        checks = @($script:checks | ForEach-Object { $_ })
        word = $WordResult
        powerPoint = $PowerPointResult
        certificate = $CertificateState
        vstoRuntime = $VstoRuntime
        msi = $MsiState
        oleLocalServer = $OleState
        companion = $CompanionState
        wordDisabledItems = $WordDisabledItems
        powerPointDisabledItems = $PowerPointDisabledItems
        recentLoadEvents = $Events
        failures = @($script:failures | ForEach-Object { $_ })
    }
    $report | ConvertTo-Json -Depth 10 | Set-Content -LiteralPath $ReportPath -Encoding UTF8
    Write-Host "Runtime diagnostic report: $ReportPath"
}

$wordResult = $null
$powerPointResult = $null
$certificateState = $null
$vstoRuntime = $null
$msiState = $null
$oleState = $null
$companionState = $null
$wordDisabledItems = @()
$powerPointDisabledItems = @()
$events = @()

try {
    $running = @(Get-Process WINWORD, POWERPNT -ErrorAction SilentlyContinue)
    if ($running.Count -gt 0) {
        throw "Word and PowerPoint must both be closed before runtime verification. Running: $($running.ProcessName -join ', ')"
    }
    Add-Check "Word and PowerPoint closed" $true "No existing WINWORD.EXE or POWERPNT.EXE process was found."

    $script:resolvedOfficePlatform = Resolve-OfficePlatform $OfficePlatform
    Add-Check "Office platform" $true $script:resolvedOfficePlatform

    $modeKey = "HKCU:\Software\VisualTeX\OfficeIntegration"
    $installationState = Get-ItemProperty -LiteralPath $modeKey -ErrorAction SilentlyContinue
    $filesAndRegistryVerified = $null -ne $installationState -and
        [int]$installationState.FilesAndRegistryVerified -eq 1 -and
        [string]$installationState.Mode -eq "vsto"
    Add-Check "Files and registry installation stage" $filesAndRegistryVerified "Mode=$($installationState.Mode); FilesAndRegistryVerified=$($installationState.FilesAndRegistryVerified)"

    $certificateState = Test-CertificateTrusted
    Add-Check "Current-user HTTPS certificate" ([bool]$certificateState.passed) ([string]$certificateState.details)

    $companionState = Test-CompanionRuntime $installationState $VisualTeXPath
    $portOwnerPathMatches =
        -not [string]::IsNullOrWhiteSpace([string]$companionState.portOwner.processPath) -and
        [string]::Equals(
            [IO.Path]::GetFullPath([string]$companionState.portOwner.processPath),
            [IO.Path]::GetFullPath([string]$companionState.executable),
            [StringComparison]::OrdinalIgnoreCase)
    $portOwnerPidMatchesVerifiedProcess =
        $null -ne $companionState.processId -and
        $null -ne $companionState.portOwner.processId -and
        [int]$companionState.processId -eq [int]$companionState.portOwner.processId -and
        [bool]$companionState.processRunning
    $portOwnedByVisualTeX =
        [bool]$companionState.portListening -and
        ($portOwnerPathMatches -or $portOwnerPidMatchesVerifiedProcess)
    Add-Check "Companion process" ([bool]$companionState.processRunning) "Executable=$($companionState.executable); PID=$($companionState.processId); StartedPID=$($companionState.startedProcessId); ExitCode=$($companionState.startedProcessExitCode); AppDataRoot=$($companionState.appDataRoot)"
    Add-Check "Companion TCP port" $portOwnedByVisualTeX "127.0.0.1:$($companionState.port); Owner=$($companionState.portOwner.processName) PID=$($companionState.portOwner.processId) Path=$($companionState.portOwner.processPath)"
    Add-Check "Companion HTTPS health" ([bool]$companionState.httpsHealthy) $(if ($companionState.httpsHealthy) { "ok=true; Raw=$($companionState.healthRaw)" } else { "Stage=$($companionState.stage); Errors=$($companionState.errors -join '; '); TLS=$($companionState.tlsPolicyErrors); Exceptions=$($companionState.exceptionChain | ConvertTo-Json -Depth 5 -Compress)" })
    if ([string]$companionState.stage -eq "port-conflict") {
        Write-Host "[SKIP] Companion certificate and protocol - port ownership must be corrected before TLS validation."
    } else {
        Add-Check "Companion certificate" ([bool]$companionState.certificateMatches) "Registry=$($companionState.expectedThumbprint); File=$($companionState.localThumbprint); Server=$($companionState.remoteThumbprint); TLS=$($companionState.tlsPolicyErrors)"
        Add-Check "Companion protocol" ([bool]$companionState.protocolMatches) "Expected=$($companionState.expectedProtocol); Actual=$($companionState.health.protocolVersion); Response=$($companionState.healthRaw)"
    }

    $vstoRuntime = Get-VstoRuntimeState $script:resolvedOfficePlatform
    Add-Check "VSTO Runtime" ([bool]$vstoRuntime.installed) "$($vstoRuntime.key); Install=$($vstoRuntime.install); VSTORFeature_CLR40=$($vstoRuntime.clr40); Version=$($vstoRuntime.version); registryView=$($vstoRuntime.registryView)"

    $msiState = Test-MsiInstalled 10
    Add-Check "Office MSI installed" ([bool]$msiState.installed) "ProductCode=$($msiState.productCode); Source=$($msiState.source)"

    $wordLoadBehaviorBefore = Get-LoadBehavior Word $script:resolvedOfficePlatform
    $powerPointLoadBehaviorBefore = Get-LoadBehavior PowerPoint $script:resolvedOfficePlatform
    Add-Check "Word LoadBehavior before startup" ($wordLoadBehaviorBefore -eq 3) "LoadBehavior=$wordLoadBehaviorBefore"
    Add-Check "PowerPoint LoadBehavior before startup" ($powerPointLoadBehaviorBefore -eq 3) "LoadBehavior=$powerPointLoadBehaviorBefore"

    $oleState = Test-OleLocalServer $script:resolvedOfficePlatform
    Add-Check "OLE LocalServer" ([bool]$oleState.healthy) "CLSID=$($oleState.clsid); Server=$($oleState.server); EmbeddingProbe=$($oleState.embeddingProbe); Error=$($oleState.error)"

    if ($CompanionOnly) {
        $wordConnected = $null -ne $installationState -and [int]$installationState.WordConnected -eq 1
        $powerPointConnected = $null -ne $installationState -and [int]$installationState.PowerPointConnected -eq 1
        $wordResult = [pscustomobject]@{
            host = "Word"
            progId = "VisualTeX.WordVsto"
            connected = $wordConnected
            skipped = $true
            description = "COMAddIn.Connect was not tested during companion-only installation verification."
            error = ""
        }
        $powerPointResult = [pscustomobject]@{
            host = "PowerPoint"
            progId = "VisualTeX.PowerPointVsto"
            connected = $powerPointConnected
            skipped = $true
            description = "COMAddIn.Connect was not tested during companion-only installation verification."
            error = ""
        }
        Write-Host "[SKIP] Word/PowerPoint COMAddIn.Connect - companion-only verification does not launch Office."
    } else {
        $wordResult = Test-OfficeComAddIn Word "VisualTeX.WordVsto"
        Add-Check "Word COMAddIn.Connect" ([bool]$wordResult.connected) $(if ($wordResult.connected) { "True" } else { "False; $($wordResult.error)" })
        $wordLoadBehaviorAfter = Get-LoadBehavior Word $script:resolvedOfficePlatform
        Add-Check "Word LoadBehavior after startup" ($wordLoadBehaviorAfter -eq 3) "Before=$wordLoadBehaviorBefore; After=$wordLoadBehaviorAfter"

        $powerPointResult = Test-OfficeComAddIn PowerPoint "VisualTeX.PowerPointVsto"
        Add-Check "PowerPoint COMAddIn.Connect" ([bool]$powerPointResult.connected) $(if ($powerPointResult.connected) { "True" } else { "False; $($powerPointResult.error)" })
        $powerPointLoadBehaviorAfter = Get-LoadBehavior PowerPoint $script:resolvedOfficePlatform
        Add-Check "PowerPoint LoadBehavior after startup" ($powerPointLoadBehaviorAfter -eq 3) "Before=$powerPointLoadBehaviorBefore; After=$powerPointLoadBehaviorAfter"

        if (-not $wordResult.connected) {
            $wordDisabledItems = @(Get-DisabledItems Word)
            Write-Warning "Word add-in did not connect. LoadBehavior after startup=$wordLoadBehaviorAfter; DisabledItems=$($wordDisabledItems | ConvertTo-Json -Depth 5 -Compress)"
        }
        if (-not $powerPointResult.connected) {
            $powerPointDisabledItems = @(Get-DisabledItems PowerPoint)
            Write-Warning "PowerPoint add-in did not connect. LoadBehavior after startup=$powerPointLoadBehaviorAfter; DisabledItems=$($powerPointDisabledItems | ConvertTo-Json -Depth 5 -Compress)"
        }
    }
    if ($script:failures.Count -gt 0) {
        $events = @(Get-RecentOfficeLoadEvents)
        Write-Warning "Recent Office/VSTO/.NET load events: $($events | ConvertTo-Json -Depth 5 -Compress)"
    }

    if (-not (Test-Path -LiteralPath $modeKey)) { New-Item -Path $modeKey -Force | Out-Null }
    if (-not $CompanionOnly) {
        New-ItemProperty -LiteralPath $modeKey -Name "WordConnected" -PropertyType DWord -Value ([int][bool]$wordResult.connected) -Force | Out-Null
        New-ItemProperty -LiteralPath $modeKey -Name "PowerPointConnected" -PropertyType DWord -Value ([int][bool]$powerPointResult.connected) -Force | Out-Null
    }
    $companionVerified =
        [bool]$companionState.processRunning -and
        [bool]$companionState.portListening -and
        [bool]$companionState.httpsHealthy -and
        [bool]$companionState.certificateMatches -and
        [bool]$companionState.protocolMatches
    New-ItemProperty -LiteralPath $modeKey -Name "CompanionProcessRunning" -PropertyType DWord -Value ([int][bool]$companionState.processRunning) -Force | Out-Null
    New-ItemProperty -LiteralPath $modeKey -Name "CompanionPortListening" -PropertyType DWord -Value ([int][bool]$companionState.portListening) -Force | Out-Null
    New-ItemProperty -LiteralPath $modeKey -Name "CompanionHttpsHealthy" -PropertyType DWord -Value ([int][bool]$companionState.httpsHealthy) -Force | Out-Null
    New-ItemProperty -LiteralPath $modeKey -Name "CompanionCertificateMatches" -PropertyType DWord -Value ([int][bool]$companionState.certificateMatches) -Force | Out-Null
    New-ItemProperty -LiteralPath $modeKey -Name "CompanionProtocolMatches" -PropertyType DWord -Value ([int][bool]$companionState.protocolMatches) -Force | Out-Null
    New-ItemProperty -LiteralPath $modeKey -Name "OfficeRuntimeVerified" -PropertyType DWord -Value ([int]$companionVerified) -Force | Out-Null
    New-ItemProperty -LiteralPath $modeKey -Name "LastRuntimeError" -PropertyType String -Value $(if ($script:failures.Count -eq 0) { "" } else { $script:failures -join "; " }) -Force | Out-Null
    New-ItemProperty -LiteralPath $modeKey -Name "LastRuntimeReport" -PropertyType String -Value $ReportPath -Force | Out-Null

    Write-Report $wordResult $powerPointResult $certificateState $vstoRuntime $msiState $oleState $companionState $wordDisabledItems $powerPointDisabledItems $events
    if ($script:failures.Count -gt 0) {
        throw "Office runtime verification failed: $($script:failures -join '; ')"
    }

    if ($CompanionOnly) {
        Write-Host "installed and runtime-verified: static Office installation and VisualTeX companion runtime passed."
        Write-Warning "Word/PowerPoint COMAddIn.Connect has not been tested because Office applications were not launched."
    } else {
        Write-Host "Native Office integration installed and verified successfully."
    }
} catch {
    if ($script:failures.Count -eq 0) {
        Add-Check "Runtime verification" $false $_.Exception.Message
    }
    if ($null -eq $wordResult) {
        $wordResult = [pscustomobject]@{ host = "Word"; progId = "VisualTeX.WordVsto"; connected = $false; description = ""; error = $_.Exception.Message }
    }
    if ($null -eq $powerPointResult) {
        $powerPointResult = [pscustomobject]@{ host = "PowerPoint"; progId = "VisualTeX.PowerPointVsto"; connected = $false; description = ""; error = $_.Exception.Message }
    }
    if ($null -eq $certificateState) { $certificateState = [pscustomobject]@{ passed = $false; details = "Not evaluated" } }
    if ($null -eq $vstoRuntime) { $vstoRuntime = [pscustomobject]@{ installed = $false; install = $null; clr40 = $null; version = $null; registryView = "none"; key = "Not evaluated" } }
    if ($null -eq $msiState) { $msiState = [pscustomobject]@{ installed = $false; productCode = "" } }
    if ($null -eq $oleState) { $oleState = [pscustomobject]@{ healthy = $false; server = ""; clsid = "{8FF7F5AA-0D60-48D5-ADBD-65A64B4C827B}"; embeddingProbe = $false; error = "Not evaluated" } }
    if ($null -eq $companionState) { $companionState = [pscustomobject]@{ executable = ""; appDataRoot = ""; port = 0; expectedProtocol = 0; processRunning = $false; portListening = $false; httpsHealthy = $false; certificateMatches = $false; protocolMatches = $false; certificatePath = ""; expectedThumbprint = ""; localThumbprint = ""; remoteThumbprint = ""; health = $null; errors = @("Not evaluated") } }
    try {
        $modeKey = "HKCU:\Software\VisualTeX\OfficeIntegration"
        if (-not (Test-Path -LiteralPath $modeKey)) { New-Item -Path $modeKey -Force | Out-Null }
        $failedRuntimeValues = @(
            "CompanionProcessRunning",
            "CompanionPortListening",
            "CompanionHttpsHealthy",
            "CompanionCertificateMatches",
            "CompanionProtocolMatches",
            "OfficeRuntimeVerified"
        )
        if (-not $CompanionOnly) {
            $failedRuntimeValues += @("WordConnected", "PowerPointConnected")
        }
        foreach ($valueName in $failedRuntimeValues) {
            New-ItemProperty -LiteralPath $modeKey -Name $valueName -PropertyType DWord -Value 0 -Force | Out-Null
        }
        $runtimeError = if ($script:failures.Count -gt 0) { $script:failures -join "; " } else { $_.Exception.Message }
        New-ItemProperty -LiteralPath $modeKey -Name "LastRuntimeError" -PropertyType String -Value $runtimeError -Force | Out-Null
        New-ItemProperty -LiteralPath $modeKey -Name "LastRuntimeReport" -PropertyType String -Value $ReportPath -Force | Out-Null
        Write-Report $wordResult $powerPointResult $certificateState $vstoRuntime $msiState $oleState $companionState $wordDisabledItems $powerPointDisabledItems $events
    } catch { }
    throw
}
