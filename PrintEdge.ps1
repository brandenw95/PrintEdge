[CmdletBinding()]
param(
    [string]$ConfigPath = "",
    [switch]$ValidateConfig,
    [switch]$Once,
    [switch]$NoTray,
    [switch]$ShowConsole,
    [switch]$WhatIfMode
)

if ([string]::IsNullOrWhiteSpace($ConfigPath)) {
    $ConfigPath = Join-Path -Path $PSScriptRoot -ChildPath "PrintEdge.config.json"
}

Set-StrictMode -Version 2.0
$ErrorActionPreference = "Stop"

$script:AppName = "PrintEdge"
$script:ManagedTag = "ManagedBy=PrintEdge"
$script:Config = $null
$script:CurrentProfile = $null
$script:LastIPv4Address = $null
$script:SyncTaskPath = "\PrintEdge\Sync"

function Get-PropertyValue {
    param(
        [object]$InputObject,
        [Parameter(Mandatory)]
        [string]$Name,
        [object]$Default = $null
    )

    if ($null -eq $InputObject) {
        return $Default
    }

    $property = $InputObject.PSObject.Properties[$Name]
    if ($null -eq $property -or $null -eq $property.Value) {
        return $Default
    }

    if ($property.Value -is [string] -and [string]::IsNullOrWhiteSpace($property.Value)) {
        return $Default
    }

    return $property.Value
}

function Get-ConfigArray {
    param(
        [object]$InputObject,
        [Parameter(Mandatory)]
        [string]$Name
    )

    $value = Get-PropertyValue -InputObject $InputObject -Name $Name -Default @()
    if ($null -eq $value) {
        return @()
    }

    return @($value)
}

function Get-FirstPropertyValue {
    param(
        [object[]]$Objects,
        [string[]]$Names,
        [object]$Default = $null
    )

    foreach ($item in $Objects) {
        foreach ($name in $Names) {
            $value = Get-PropertyValue -InputObject $item -Name $name -Default $null
            if ($null -ne $value) {
                return $value
            }
        }
    }

    return $Default
}

function ConvertTo-Bool {
    param(
        [object]$Value,
        [bool]$Default = $false
    )

    if ($null -eq $Value) {
        return $Default
    }

    if ($Value -is [bool]) {
        return $Value
    }

    return [System.Convert]::ToBoolean($Value)
}

function Expand-PrintEdgePath {
    param(
        [string]$Path,
        [string]$BasePath = $PSScriptRoot
    )

    if ([string]::IsNullOrWhiteSpace($Path)) {
        return $null
    }

    $expandedPath = [Environment]::ExpandEnvironmentVariables($Path)
    if ([System.IO.Path]::IsPathRooted($expandedPath)) {
        return $expandedPath
    }

    return (Join-Path -Path $BasePath -ChildPath $expandedPath)
}

function ConvertTo-SafePathSegment {
    param(
        [string]$Value,
        [string]$Fallback = "Driver"
    )

    $segment = $Value
    if ([string]::IsNullOrWhiteSpace($segment)) {
        $segment = $Fallback
    }

    foreach ($invalidChar in [System.IO.Path]::GetInvalidFileNameChars()) {
        $segment = $segment.Replace([string]$invalidChar, "_")
    }

    $segment = ($segment -replace "\s+", " ").Trim().TrimEnd(".")
    if ([string]::IsNullOrWhiteSpace($segment)) {
        return $Fallback
    }

    return $segment
}

function Get-Setting {
    param(
        [Parameter(Mandatory)]
        [string]$Name,
        [object]$Default = $null
    )

    if ($null -eq $script:Config) {
        return $Default
    }

    $settings = Get-PropertyValue -InputObject $script:Config -Name "settings" -Default $null
    return (Get-PropertyValue -InputObject $settings -Name $Name -Default $Default)
}

function Write-PrintEdgeLog {
    param(
        [Parameter(Mandatory)]
        [string]$Message,
        [ValidateSet("INFO", "WARN", "ERROR")]
        [string]$Level = "INFO"
    )

    $line = "{0} [{1}] {2}" -f (Get-Date -Format "yyyy-MM-dd HH:mm:ss"), $Level, $Message
    Write-Host $line

    $configuredLogPath = Get-Setting -Name "logPath" -Default $null
    $logPath = Expand-PrintEdgePath -Path $configuredLogPath
    if ([string]::IsNullOrWhiteSpace($logPath)) {
        return
    }

    try {
        $logDirectory = Split-Path -Path $logPath -Parent
        if (-not [string]::IsNullOrWhiteSpace($logDirectory) -and -not (Test-Path -Path $logDirectory)) {
            New-Item -Path $logDirectory -ItemType Directory -Force | Out-Null
        }

        Add-Content -Path $logPath -Value $line
    }
    catch {
        Write-Verbose ("Unable to write PrintEdge log: {0}" -f $_.Exception.Message)
    }
}

function Get-PrintEdgeLogPath {
    $configuredLogPath = Get-Setting -Name "logPath" -Default $null
    return (Expand-PrintEdgePath -Path $configuredLogPath)
}

function Get-RecentPrintEdgeLogLines {
    param(
        [int]$LineCount = 80
    )

    $logPath = Get-PrintEdgeLogPath
    if ([string]::IsNullOrWhiteSpace($logPath)) {
        return @("Log path is not configured.")
    }

    if (-not (Test-Path -Path $logPath)) {
        return @("Log file has not been created yet: {0}" -f $logPath)
    }

    try {
        return @(Get-Content -Path $logPath -Tail $LineCount -ErrorAction Stop)
    }
    catch {
        return @("Unable to read log file '{0}': {1}" -f $logPath, $_.Exception.Message)
    }
}

function Test-IsAdministrator {
    $identity = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object Security.Principal.WindowsPrincipal($identity)
    return $principal.IsInRole([Security.Principal.WindowsBuiltinRole]::Administrator)
}

function Test-PrintEdgeScheduledSyncAvailable {
    try {
        $process = Start-Process -FilePath "schtasks.exe" -ArgumentList @("/Query", "/TN", $script:SyncTaskPath) -WindowStyle Hidden -Wait -PassThru -ErrorAction Stop
        return ($process.ExitCode -eq 0)
    }
    catch {
        return $false
    }
}

function Start-PrintEdgeScheduledSync {
    param(
        [string]$Reason = "sync requested"
    )

    if (-not (Test-PrintEdgeScheduledSyncAvailable)) {
        Write-PrintEdgeLog -Level "WARN" -Message ("Elevated sync task '{0}' is not installed; running in current user context." -f $script:SyncTaskPath)
        return $false
    }

    try {
        $process = Start-Process -FilePath "schtasks.exe" -ArgumentList @("/Run", "/TN", $script:SyncTaskPath) -WindowStyle Hidden -Wait -PassThru -ErrorAction Stop
        if ($process.ExitCode -eq 0) {
            Write-PrintEdgeLog -Message ("Requested elevated sync task '{0}' because {1}." -f $script:SyncTaskPath, $Reason)
            return $true
        }

        Write-PrintEdgeLog -Level "WARN" -Message ("Unable to request elevated sync task '{0}'. schtasks exit code: {1}" -f $script:SyncTaskPath, $process.ExitCode)
    }
    catch {
        Write-PrintEdgeLog -Level "WARN" -Message ("Unable to request elevated sync task '{0}': {1}" -f $script:SyncTaskPath, $_.Exception.Message)
    }

    return $false
}

function Get-PrintEdgeDiagnosticText {
    param(
        [object]$Profile,
        [string]$CurrentIPv4Address
    )

    $profileName = "No matching profile"
    $profileCidr = "n/a"
    if ($Profile) {
        $profileName = Get-PropertyValue -InputObject $Profile -Name "name" -Default "Unnamed profile"
        $profileCidr = Get-PropertyValue -InputObject $Profile -Name "cidr" -Default "Unknown subnet"
    }

    $adminText = "No"
    if (Test-IsAdministrator) {
        $adminText = "Yes"
    }

    $scheduledSyncText = "Not installed"
    if (Test-PrintEdgeScheduledSyncAvailable) {
        $scheduledSyncText = "Installed"
    }

    $headerLines = @(
        "PrintEdge diagnostics",
        "Active IPv4: {0}" -f $CurrentIPv4Address,
        "Matched profile: {0} ({1})" -f $profileName, $profileCidr,
        "Running as admin: {0}" -f $adminText,
        "Elevated sync task: {0}" -f $scheduledSyncText,
        "Last synced IPv4: {0}" -f $script:LastIPv4Address,
        "Log: {0}" -f (Get-PrintEdgeLogPath),
        "",
        "Recent activity"
    )

    $logLines = @(Get-RecentPrintEdgeLogLines -LineCount 80)
    return (($headerLines + $logLines) -join [Environment]::NewLine)
}

function Initialize-NativeMethods {
    if ("PrintEdge.NativeMethods" -as [type]) {
        return
    }

    Add-Type -Namespace PrintEdge -Name NativeMethods -MemberDefinition @"
[DllImport("kernel32.dll")]
public static extern System.IntPtr GetConsoleWindow();

[DllImport("user32.dll")]
public static extern bool ShowWindow(System.IntPtr hWnd, int nCmdShow);

[DllImport("user32.dll")]
public static extern bool DestroyIcon(System.IntPtr hIcon);
"@
}

function Hide-ConsoleWindow {
    Initialize-NativeMethods
    $consoleHandle = [PrintEdge.NativeMethods]::GetConsoleWindow()
    if ($consoleHandle -ne [IntPtr]::Zero) {
        [PrintEdge.NativeMethods]::ShowWindow($consoleHandle, 0) | Out-Null
    }
}

function New-PrintEdgeTrayIcon {
    Initialize-NativeMethods

    $bitmap = New-Object System.Drawing.Bitmap(32, 32)
    $graphics = [System.Drawing.Graphics]::FromImage($bitmap)
    $graphics.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
    $graphics.Clear([System.Drawing.Color]::Transparent)

    $backgroundBrush = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::FromArgb(24, 119, 242))
    $paperBrush = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::FromArgb(245, 248, 255))
    $bodyBrush = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::FromArgb(14, 55, 116))
    $accentBrush = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::FromArgb(87, 219, 172))
    $whitePen = New-Object System.Drawing.Pen([System.Drawing.Color]::White, 2)

    try {
        $graphics.FillEllipse($backgroundBrush, 1, 1, 30, 30)
        $graphics.FillRectangle($paperBrush, 10, 6, 12, 8)
        $graphics.DrawRectangle($whitePen, 10, 6, 12, 8)
        $graphics.FillRectangle($bodyBrush, 7, 13, 18, 10)
        $graphics.FillRectangle($paperBrush, 10, 20, 12, 6)
        $graphics.FillEllipse($accentBrush, 20, 15, 3, 3)
        $graphics.DrawLine($whitePen, 11, 24, 21, 24)

        $iconHandle = $bitmap.GetHicon()
        try {
            $icon = [System.Drawing.Icon]::FromHandle($iconHandle)
            return $icon.Clone()
        }
        finally {
            [PrintEdge.NativeMethods]::DestroyIcon($iconHandle) | Out-Null
        }
    }
    finally {
        $graphics.Dispose()
        $backgroundBrush.Dispose()
        $paperBrush.Dispose()
        $bodyBrush.Dispose()
        $accentBrush.Dispose()
        $whitePen.Dispose()
        $bitmap.Dispose()
    }
}

function Convert-IPv4ToUInt32 {
    param(
        [Parameter(Mandatory)]
        [string]$Address
    )

    $bytes = [System.Net.IPAddress]::Parse($Address).GetAddressBytes()
    if ($bytes.Count -ne 4) {
        throw "'$Address' is not an IPv4 address."
    }

    return [uint32](
        ([uint64]$bytes[0] -shl 24) -bor
        ([uint64]$bytes[1] -shl 16) -bor
        ([uint64]$bytes[2] -shl 8) -bor
        [uint64]$bytes[3]
    )
}

function Test-IPv4InCidr {
    param(
        [Parameter(Mandatory)]
        [string]$IPAddress,
        [Parameter(Mandatory)]
        [string]$Cidr
    )

    $parts = $Cidr -split "/"
    if ($parts.Count -ne 2) {
        throw "CIDR '$Cidr' must be in the form 192.168.1.0/24."
    }

    $prefixLength = [int]$parts[1]
    if ($prefixLength -lt 0 -or $prefixLength -gt 32) {
        throw "CIDR '$Cidr' has an invalid prefix length."
    }

    $ip = Convert-IPv4ToUInt32 -Address $IPAddress
    $network = Convert-IPv4ToUInt32 -Address $parts[0]

    if ($prefixLength -eq 0) {
        return $true
    }

    $mask = [uint32](([uint64]4294967295 -shl (32 - $prefixLength)) -band [uint64]4294967295)
    return (($ip -band $mask) -eq ($network -band $mask))
}

function Get-ActiveIPv4Address {
    try {
        $activeInterface = Get-NetIPConfiguration |
            Where-Object { $null -ne $_.IPv4DefaultGateway -and $null -ne $_.IPv4Address } |
            Select-Object -First 1

        if ($activeInterface -and $activeInterface.IPv4Address) {
            return $activeInterface.IPv4Address.IPAddress
        }
    }
    catch {
        Write-PrintEdgeLog -Level "WARN" -Message ("Unable to query Get-NetIPConfiguration: {0}" -f $_.Exception.Message)
    }

    try {
        $ipAddress = Get-NetIPAddress -AddressFamily IPv4 |
            Where-Object {
                $_.IPAddress -notlike "169.254.*" -and
                $_.IPAddress -ne "127.0.0.1" -and
                $_.PrefixOrigin -ne "WellKnown"
            } |
            Select-Object -First 1

        if ($ipAddress) {
            return $ipAddress.IPAddress
        }
    }
    catch {
        Write-PrintEdgeLog -Level "WARN" -Message ("Unable to query Get-NetIPAddress: {0}" -f $_.Exception.Message)
    }

    return $null
}

function GetSubnet {
    return (Get-ActiveIPv4Address)
}

function Import-PrintEdgeConfig {
    param(
        [Parameter(Mandatory)]
        [string]$Path
    )

    $resolvedPath = Expand-PrintEdgePath -Path $Path
    if (-not (Test-Path -Path $resolvedPath)) {
        throw "Config file not found: $resolvedPath"
    }

    $config = Get-Content -Path $resolvedPath -Raw | ConvertFrom-Json
    $script:Config = $config

    Test-PrintEdgeConfig -Config $config | Out-Null
    return $config
}

function Test-PrintEdgeConfig {
    param(
        [Parameter(Mandatory)]
        [object]$Config
    )

    $subnets = @(Get-ConfigArray -InputObject $Config -Name "subnets")
    if ($subnets.Count -eq 0) {
        throw "Config must contain at least one subnet profile."
    }

    foreach ($profile in $subnets) {
        $profileName = Get-PropertyValue -InputObject $profile -Name "name" -Default "<unnamed>"
        $cidr = Get-PropertyValue -InputObject $profile -Name "cidr" -Default $null
        if ([string]::IsNullOrWhiteSpace($cidr)) {
            throw "Subnet profile '$profileName' is missing cidr."
        }

        Test-IPv4InCidr -IPAddress (($cidr -split "/")[0]) -Cidr $cidr | Out-Null

        $printerMap = @{}
        foreach ($printer in @(Get-ConfigArray -InputObject $profile -Name "printers")) {
            $printerId = Get-PropertyValue -InputObject $printer -Name "id" -Default $null
            $hostAddress = Get-PropertyValue -InputObject $printer -Name "hostAddress" -Default $null

            if ([string]::IsNullOrWhiteSpace($printerId)) {
                throw "Subnet profile '$profileName' has a printer without an id."
            }

            if ([string]::IsNullOrWhiteSpace($hostAddress)) {
                throw "Printer '$printerId' in subnet profile '$profileName' is missing hostAddress."
            }

            Convert-IPv4ToUInt32 -Address $hostAddress | Out-Null
            $printerMap[$printerId] = $printer
        }

        $queues = @(Get-ConfigArray -InputObject $profile -Name "printQueues")
        if ($queues.Count -eq 0) {
            throw "Subnet profile '$profileName' must define at least one print queue."
        }

        foreach ($queue in $queues) {
            $queueName = Get-FirstPropertyValue -Objects @($queue) -Names @("name", "queueName", "displayName") -Default $null
            $printerId = Get-PropertyValue -InputObject $queue -Name "printerId" -Default $null

            if ([string]::IsNullOrWhiteSpace($queueName)) {
                throw "Subnet profile '$profileName' contains a queue without a name."
            }

            if ([string]::IsNullOrWhiteSpace($printerId) -or -not $printerMap.ContainsKey($printerId)) {
                throw "Queue '$queueName' in subnet profile '$profileName' references unknown printerId '$printerId'."
            }

            $driverName = Get-FirstPropertyValue -Objects @($queue, $printerMap[$printerId]) -Names @("driverName", "defaultDriverName") -Default $null
            $connectionName = Get-FirstPropertyValue -Objects @($queue, $printerMap[$printerId]) -Names @("connectionName", "sharePath", "uncPath") -Default $null
            if ([string]::IsNullOrWhiteSpace($driverName) -and [string]::IsNullOrWhiteSpace($connectionName)) {
                throw "Queue '$queueName' in subnet profile '$profileName' needs either connectionName for non-admin per-user mapping or driverName/defaultDriverName for direct local queue mode."
            }
        }
    }

    return $true
}

function Select-SubnetProfile {
    param(
        [Parameter(Mandatory)]
        [object]$Config,
        [Parameter(Mandatory)]
        [string]$IPAddress
    )

    foreach ($profile in @(Get-ConfigArray -InputObject $Config -Name "subnets")) {
        $cidr = Get-PropertyValue -InputObject $profile -Name "cidr" -Default $null
        if (-not [string]::IsNullOrWhiteSpace($cidr) -and (Test-IPv4InCidr -IPAddress $IPAddress -Cidr $cidr)) {
            return $profile
        }
    }

    return $null
}

function Get-PrinterMap {
    param(
        [Parameter(Mandatory)]
        [object]$Profile
    )

    $printerMap = @{}
    foreach ($printer in @(Get-ConfigArray -InputObject $Profile -Name "printers")) {
        $printerId = Get-PropertyValue -InputObject $printer -Name "id" -Default $null
        if (-not [string]::IsNullOrWhiteSpace($printerId)) {
            $printerMap[$printerId] = $printer
        }
    }

    return $printerMap
}

function Get-DesiredPrintQueues {
    param(
        [Parameter(Mandatory)]
        [object]$Profile
    )

    $printerMap = Get-PrinterMap -Profile $Profile
    $profileName = Get-PropertyValue -InputObject $Profile -Name "name" -Default "Unknown"
    $prefix = [string](Get-Setting -Name "managedNamePrefix" -Default "PrintEdge - ")
    $prefixQueueNames = ConvertTo-Bool -Value (Get-Setting -Name "prefixQueueNames" -Default $true) -Default $true
    $desiredQueues = @()

    foreach ($queue in @(Get-ConfigArray -InputObject $Profile -Name "printQueues")) {
        $printerId = Get-PropertyValue -InputObject $queue -Name "printerId" -Default $null
        $printer = $printerMap[$printerId]
        $hostAddress = Get-FirstPropertyValue -Objects @($queue, $printer) -Names @("hostAddress") -Default $null
        $queueName = Get-FirstPropertyValue -Objects @($queue) -Names @("name", "queueName", "displayName") -Default $null
        $printerDisplayName = Get-FirstPropertyValue -Objects @($printer, $queue) -Names @("displayName", "name") -Default $printerId
        $driverName = Get-FirstPropertyValue -Objects @($queue, $printer) -Names @("driverName", "defaultDriverName") -Default $null
        $connectionName = Get-FirstPropertyValue -Objects @($queue, $printer) -Names @("connectionName", "sharePath", "uncPath") -Default $null
        $location = Get-FirstPropertyValue -Objects @($queue, $printer, $Profile) -Names @("location", "name") -Default $profileName
        $comment = Get-FirstPropertyValue -Objects @($queue, $printer) -Names @("comment", "description") -Default "Managed print queue"
        $portName = Get-FirstPropertyValue -Objects @($queue) -Names @("portName") -Default $null

        if ([string]::IsNullOrWhiteSpace($portName)) {
            $portName = "PE_{0}" -f $hostAddress
        }

        if ($prefixQueueNames -and -not $queueName.StartsWith($prefix, [System.StringComparison]::OrdinalIgnoreCase)) {
            $queueName = "{0}{1}" -f $prefix, $queueName
        }

        $desiredQueues += [pscustomobject]@{
            Name = $queueName
            PrinterId = $printerId
            PrinterDisplayName = $printerDisplayName
            ConnectionName = $connectionName
            HostAddress = $hostAddress
            PortName = $portName
            DriverName = $driverName
            InfPath = Get-FirstPropertyValue -Objects @($queue, $printer) -Names @("infPath") -Default $null
            DriverBlobPath = Get-FirstPropertyValue -Objects @($queue, $printer) -Names @("driverBlobPath", "azureBlobPath") -Default $null
            DriverUrl = Get-FirstPropertyValue -Objects @($queue, $printer) -Names @("driverUrl", "azureUrl") -Default $null
            InfRelativePath = Get-FirstPropertyValue -Objects @($queue, $printer) -Names @("infRelativePath") -Default $null
            Location = $location
            Comment = $comment
            ProfileName = $profileName
        }
    }

    return $desiredQueues
}

function Get-PrintEdgeQueueIdentityNames {
    param(
        [object[]]$DesiredQueues
    )

    $names = @()
    foreach ($queue in @($DesiredQueues)) {
        if ($null -eq $queue) {
            continue
        }

        if (-not [string]::IsNullOrWhiteSpace($queue.Name)) {
            $names += $queue.Name
        }

        if (-not [string]::IsNullOrWhiteSpace($queue.ConnectionName)) {
            $names += $queue.ConnectionName
            $names += (Split-Path -Path $queue.ConnectionName -Leaf)
        }
    }

    return @($names | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Select-Object -Unique)
}

function Get-AllConfiguredPrintEdgeQueueIdentityNames {
    $names = @()
    if ($null -eq $script:Config) {
        return @()
    }

    foreach ($profile in @(Get-ConfigArray -InputObject $script:Config -Name "subnets")) {
        $names += Get-PrintEdgeQueueIdentityNames -DesiredQueues @(Get-DesiredPrintQueues -Profile $profile)
    }

    return @($names | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Select-Object -Unique)
}

function Test-PrinterNameMatchesIdentity {
    param(
        [Parameter(Mandatory)]
        [string]$PrinterName,
        [Parameter(Mandatory)]
        [hashtable]$IdentityNames
    )

    if ($IdentityNames.ContainsKey($PrinterName)) {
        return $true
    }

    foreach ($identityName in $IdentityNames.Keys) {
        if ($identityName.Length -lt 4 -or $identityName.StartsWith("\\")) {
            continue
        }

        $escapedIdentityName = [System.Management.Automation.WildcardPattern]::Escape($identityName)
        if ($PrinterName -like "*$escapedIdentityName*") {
            return $true
        }
    }

    return $false
}

function Get-AzureDriverUri {
    param(
        [string]$DriverBlobPath,
        [string]$DriverUrl
    )

    if (-not [string]::IsNullOrWhiteSpace($DriverUrl)) {
        return $DriverUrl
    }

    $baseUrl = [string](Get-Setting -Name "azureBlobBaseUrl" -Default "")
    if ([string]::IsNullOrWhiteSpace($baseUrl) -or [string]::IsNullOrWhiteSpace($DriverBlobPath)) {
        return $null
    }

    $segments = $DriverBlobPath -split "/" | ForEach-Object { [Uri]::EscapeDataString($_) }
    $uri = "{0}/{1}" -f $baseUrl.TrimEnd("/"), ($segments -join "/")
    $sasToken = [string](Get-Setting -Name "azureSasToken" -Default "")
    if ([string]::IsNullOrWhiteSpace($sasToken)) {
        return $uri
    }

    if ($sasToken.StartsWith("?")) {
        return "{0}{1}" -f $uri, $sasToken
    }

    return "{0}?{1}" -f $uri, $sasToken
}

function Resolve-AzureDriverPackage {
    param(
        [Parameter(Mandatory)]
        [object]$DesiredQueue
    )

    $azureEnabled = ConvertTo-Bool -Value (Get-Setting -Name "enableAzureDriverDownload" -Default $false) -Default $false
    if (-not $azureEnabled) {
        return $null
    }

    $driverUri = Get-AzureDriverUri -DriverBlobPath $DesiredQueue.DriverBlobPath -DriverUrl $DesiredQueue.DriverUrl
    if ([string]::IsNullOrWhiteSpace($driverUri)) {
        return $null
    }

    $cacheRoot = Expand-PrintEdgePath -Path ([string](Get-Setting -Name "driverCachePath" -Default "Drivers"))
    $driverFolderName = ConvertTo-SafePathSegment -Value $DesiredQueue.DriverName -Fallback $DesiredQueue.PrinterId
    $driverCachePath = Join-Path -Path $cacheRoot -ChildPath $driverFolderName
    if (-not (Test-Path -Path $driverCachePath)) {
        New-Item -Path $driverCachePath -ItemType Directory -Force | Out-Null
    }

    $uriWithoutQuery = ($driverUri -split "\?")[0]
    $fileName = [System.IO.Path]::GetFileName($uriWithoutQuery)
    if ([string]::IsNullOrWhiteSpace($fileName)) {
        $fileName = "{0}.zip" -f $DesiredQueue.PrinterId
    }

    $downloadPath = Join-Path -Path $driverCachePath -ChildPath $fileName
    if (-not (Test-Path -Path $downloadPath)) {
        Write-PrintEdgeLog -Message ("Downloading driver package for '{0}' to '{1}'." -f $DesiredQueue.DriverName, $driverCachePath)
        Invoke-WebRequest -Uri $driverUri -OutFile $downloadPath -UseBasicParsing
    }

    if ($downloadPath.EndsWith(".zip", [System.StringComparison]::OrdinalIgnoreCase)) {
        $extractPath = Join-Path -Path $driverCachePath -ChildPath "Extracted"
        if (-not (Test-Path -Path $extractPath)) {
            New-Item -Path $extractPath -ItemType Directory -Force | Out-Null
            Expand-Archive -Path $downloadPath -DestinationPath $extractPath -Force
        }

        if (-not [string]::IsNullOrWhiteSpace($DesiredQueue.InfRelativePath)) {
            $configuredInfPath = Join-Path -Path $extractPath -ChildPath $DesiredQueue.InfRelativePath
            if (Test-Path -Path $configuredInfPath) {
                return $configuredInfPath
            }
        }

        $infFile = Get-ChildItem -Path $extractPath -Filter "*.inf" -Recurse | Select-Object -First 1
        if ($infFile) {
            return $infFile.FullName
        }

        throw "Azure driver package '$downloadPath' did not contain an INF file."
    }

    if ($downloadPath.EndsWith(".inf", [System.StringComparison]::OrdinalIgnoreCase)) {
        return $downloadPath
    }

    throw "Driver package '$downloadPath' must be a .zip or .inf file."
}

function Resolve-DriverInfPath {
    param(
        [Parameter(Mandatory)]
        [object]$DesiredQueue
    )

    $localInfPath = Expand-PrintEdgePath -Path $DesiredQueue.InfPath
    if (-not [string]::IsNullOrWhiteSpace($localInfPath) -and (Test-Path -Path $localInfPath)) {
        return $localInfPath
    }

    return (Resolve-AzureDriverPackage -DesiredQueue $DesiredQueue)
}

function Ensure-PrinterDriver {
    param(
        [Parameter(Mandatory)]
        [object]$DesiredQueue,
        [switch]$DryRun
    )

    if ([string]::IsNullOrWhiteSpace($DesiredQueue.DriverName)) {
        return
    }

    if (Get-PrinterDriver -Name $DesiredQueue.DriverName -ErrorAction SilentlyContinue) {
        return
    }

    if ($DryRun) {
        Write-PrintEdgeLog -Message ("Would verify or install driver '{0}' if required." -f $DesiredQueue.DriverName)
        return
    }

    $infPath = Resolve-DriverInfPath -DesiredQueue $DesiredQueue
    if ([string]::IsNullOrWhiteSpace($infPath)) {
        throw "Driver '$($DesiredQueue.DriverName)' is not installed and no local or Azure INF was available."
    }

    Write-PrintEdgeLog -Message ("Installing driver '{0}' from '{1}'." -f $DesiredQueue.DriverName, $infPath)
    & pnputil.exe /add-driver $infPath /install | Out-Null
    if ($LASTEXITCODE -ne 0) {
        throw "pnputil failed while installing '$infPath' with exit code $LASTEXITCODE."
    }

    if (-not (Get-PrinterDriver -Name $DesiredQueue.DriverName -ErrorAction SilentlyContinue)) {
        Add-PrinterDriver -Name $DesiredQueue.DriverName -ErrorAction Stop
    }
}

function Ensure-PrinterPort {
    param(
        [Parameter(Mandatory)]
        [object]$DesiredQueue,
        [switch]$DryRun
    )

    $existingPort = Get-PrinterPort -Name $DesiredQueue.PortName -ErrorAction SilentlyContinue
    if ($existingPort) {
        $existingHostAddress = Get-PropertyValue -InputObject $existingPort -Name "PrinterHostAddress" -Default $null
        if (-not [string]::IsNullOrWhiteSpace($existingHostAddress) -and $existingHostAddress -ne $DesiredQueue.HostAddress) {
            throw "Printer port '$($DesiredQueue.PortName)' already exists but points to '$existingHostAddress' instead of '$($DesiredQueue.HostAddress)'. Rename the configured port or correct the existing port."
        }

        return
    }

    if ($DryRun) {
        Write-PrintEdgeLog -Message ("Would create TCP/IP printer port '{0}' for '{1}'." -f $DesiredQueue.PortName, $DesiredQueue.HostAddress)
        return
    }

    Write-PrintEdgeLog -Message ("Creating TCP/IP printer port '{0}' for '{1}'." -f $DesiredQueue.PortName, $DesiredQueue.HostAddress)
    Add-PrinterPort -Name $DesiredQueue.PortName -PrinterHostAddress $DesiredQueue.HostAddress -ErrorAction Stop
}

function Ensure-PrinterConnection {
    param(
        [Parameter(Mandatory)]
        [object]$DesiredQueue,
        [switch]$DryRun
    )

    $connectionName = $DesiredQueue.ConnectionName
    if ([string]::IsNullOrWhiteSpace($connectionName)) {
        return
    }

    $existingPrinter = Get-Printer -Name $connectionName -ErrorAction SilentlyContinue
    if (-not $existingPrinter) {
        $connectionLeafName = Split-Path -Path $connectionName -Leaf
        $escapedConnectionLeafName = [System.Management.Automation.WildcardPattern]::Escape($connectionLeafName)
        $existingPrinter = Get-Printer -ErrorAction SilentlyContinue |
            Where-Object { $_.Name -eq $connectionLeafName -or $_.Name -like "*$escapedConnectionLeafName*" } |
            Select-Object -First 1
    }

    if ($existingPrinter) {
        Write-PrintEdgeLog -Message ("Printer connection '{0}' is already available." -f $connectionName)
        return
    }

    if ($DryRun) {
        Write-PrintEdgeLog -Message ("Would add per-user printer connection '{0}'." -f $connectionName)
        return
    }

    Write-PrintEdgeLog -Message ("Adding per-user printer connection '{0}'." -f $connectionName)
    Add-Printer -ConnectionName $connectionName -ErrorAction Stop
}

function Ensure-PrintQueue {
    param(
        [Parameter(Mandatory)]
        [object]$DesiredQueue,
        [switch]$DryRun
    )

    $managedComment = "{0} [{1}; PrinterId={2}; Profile={3}]" -f $DesiredQueue.Comment, $script:ManagedTag, $DesiredQueue.PrinterId, $DesiredQueue.ProfileName
    $existingPrinter = Get-Printer -Name $DesiredQueue.Name -ErrorAction SilentlyContinue

    if ($DryRun) {
        if ($existingPrinter) {
            Write-PrintEdgeLog -Message ("Would update print queue '{0}'." -f $DesiredQueue.Name)
        }
        else {
            Write-PrintEdgeLog -Message ("Would create print queue '{0}'." -f $DesiredQueue.Name)
        }
        return
    }

    if ($existingPrinter) {
        Write-PrintEdgeLog -Message ("Updating print queue '{0}'." -f $DesiredQueue.Name)
        Set-Printer -Name $DesiredQueue.Name `
            -DriverName $DesiredQueue.DriverName `
            -PortName $DesiredQueue.PortName `
            -Location $DesiredQueue.Location `
            -Comment $managedComment `
            -ErrorAction Stop
        return
    }

    Write-PrintEdgeLog -Message ("Creating print queue '{0}'." -f $DesiredQueue.Name)
    Add-Printer -Name $DesiredQueue.Name `
        -DriverName $DesiredQueue.DriverName `
        -PortName $DesiredQueue.PortName `
        -Location $DesiredQueue.Location `
        -Comment $managedComment `
        -ErrorAction Stop
}

function Remove-StalePrintEdgeQueues {
    param(
        [Parameter(Mandatory)]
        [AllowEmptyCollection()]
        [object[]]$DesiredQueues,
        [switch]$DryRun
    )

    $removeStaleManagedQueues = ConvertTo-Bool -Value (Get-Setting -Name "removeStaleManagedQueues" -Default $true) -Default $true
    $removeUnmanagedPrinters = ConvertTo-Bool -Value (Get-Setting -Name "removeUnmanagedPrinters" -Default $false) -Default $false
    if (-not $removeStaleManagedQueues -and -not $removeUnmanagedPrinters) {
        return
    }

    $desiredNames = @{}
    foreach ($name in @(Get-PrintEdgeQueueIdentityNames -DesiredQueues $DesiredQueues)) {
        $desiredNames[$name] = $true
    }

    $configuredNames = @{}
    foreach ($name in @(Get-AllConfiguredPrintEdgeQueueIdentityNames)) {
        $configuredNames[$name] = $true
    }

    $prefix = [string](Get-Setting -Name "managedNamePrefix" -Default "PrintEdge - ")
    $installedPrinters = @(Get-Printer -ErrorAction Stop)
    foreach ($printer in $installedPrinters) {
        $printerName = [string]$printer.Name
        if (Test-PrinterNameMatchesIdentity -PrinterName $printerName -IdentityNames $desiredNames) {
            continue
        }

        $isManaged = (
            ($printer.Comment -like "*$script:ManagedTag*") -or
            ($printerName.StartsWith($prefix, [System.StringComparison]::OrdinalIgnoreCase)) -or
            (Test-PrinterNameMatchesIdentity -PrinterName $printerName -IdentityNames $configuredNames)
        )

        if (-not $removeUnmanagedPrinters -and -not $isManaged) {
            continue
        }

        if ($DryRun) {
            Write-PrintEdgeLog -Message ("Would remove stale print queue '{0}'." -f $printer.Name)
            continue
        }

        try {
            Write-PrintEdgeLog -Message ("Removing stale print queue '{0}'." -f $printer.Name)
            Remove-Printer -Name $printer.Name -ErrorAction Stop
        }
        catch {
            Write-PrintEdgeLog -Level "WARN" -Message ("Unable to remove stale print queue '{0}': {1}" -f $printer.Name, $_.Exception.Message)
        }
    }
}

function Sync-PrintEdgeProfile {
    param(
        [Parameter(Mandatory)]
        [object]$Profile,
        [switch]$DryRun
    )

    $profileName = Get-PropertyValue -InputObject $Profile -Name "name" -Default "Unknown"
    $desiredQueues = @(Get-DesiredPrintQueues -Profile $Profile)
    Write-PrintEdgeLog -Message ("Applying subnet profile '{0}' with {1} print queue(s)." -f $profileName, $desiredQueues.Count)

    foreach ($desiredQueue in $desiredQueues) {
        try {
            if (-not [string]::IsNullOrWhiteSpace($desiredQueue.ConnectionName)) {
                Ensure-PrinterConnection -DesiredQueue $desiredQueue -DryRun:$DryRun
                continue
            }

            if (-not $DryRun -and -not (Test-IsAdministrator)) {
                try {
                    Resolve-DriverInfPath -DesiredQueue $desiredQueue | Out-Null
                }
                catch {
                    Write-PrintEdgeLog -Level "WARN" -Message ("Driver package for '{0}' could not be prepared without elevation: {1}" -f $desiredQueue.Name, $_.Exception.Message)
                }

                Write-PrintEdgeLog -Level "WARN" -Message ("Skipping direct local queue '{0}' because Windows requires administrator rights to install local TCP/IP ports, system drivers, and local print queues. Configure connectionName for non-admin per-user mapping." -f $desiredQueue.Name)
                continue
            }

            Ensure-PrinterDriver -DesiredQueue $desiredQueue -DryRun:$DryRun
            Ensure-PrinterPort -DesiredQueue $desiredQueue -DryRun:$DryRun
            Ensure-PrintQueue -DesiredQueue $desiredQueue -DryRun:$DryRun
        }
        catch {
            Write-PrintEdgeLog -Level "ERROR" -Message ("Unable to sync queue '{0}': {1}" -f $desiredQueue.Name, $_.Exception.Message)
        }
    }

    Remove-StalePrintEdgeQueues -DesiredQueues $desiredQueues -DryRun:$DryRun
}

function Sync-CurrentSubnet {
    param(
        [switch]$Force,
        [switch]$DryRun
    )

    $currentIPv4Address = Get-ActiveIPv4Address
    if ([string]::IsNullOrWhiteSpace($currentIPv4Address)) {
        Write-PrintEdgeLog -Level "WARN" -Message "No active IPv4 address was detected."
        return
    }

    if (-not $Force -and $script:LastIPv4Address -eq $currentIPv4Address) {
        return
    }

    $previousIPv4Address = $script:LastIPv4Address
    $script:LastIPv4Address = $currentIPv4Address
    $profile = Select-SubnetProfile -Config $script:Config -IPAddress $currentIPv4Address
    $script:CurrentProfile = $profile

    if (-not (Test-IsAdministrator) -and (Start-PrintEdgeScheduledSync -Reason ("active IPv4 changed from '{0}' to '{1}'" -f $previousIPv4Address, $currentIPv4Address))) {
        return
    }

    if ($null -eq $profile) {
        Write-PrintEdgeLog -Level "WARN" -Message ("No subnet profile matched active IPv4 address '{0}'." -f $currentIPv4Address)
        Remove-StalePrintEdgeQueues -DesiredQueues @() -DryRun:$DryRun
        return
    }

    Sync-PrintEdgeProfile -Profile $profile -DryRun:$DryRun
}

function Show-PrintersDialog {
    param(
        [Parameter(Mandatory)]
        [object]$Config
    )

    $currentIPv4Address = Get-ActiveIPv4Address
    $profile = $null
    if (-not [string]::IsNullOrWhiteSpace($currentIPv4Address)) {
        $profile = Select-SubnetProfile -Config $Config -IPAddress $currentIPv4Address
    }

    $backgroundColor = [System.Drawing.ColorTranslator]::FromHtml("#F6F8FB")
    $cardColor = [System.Drawing.Color]::White
    $textColor = [System.Drawing.ColorTranslator]::FromHtml("#172033")
    $mutedColor = [System.Drawing.ColorTranslator]::FromHtml("#627084")
    $accentColor = [System.Drawing.ColorTranslator]::FromHtml("#1877F2")
    $successColor = [System.Drawing.ColorTranslator]::FromHtml("#148A5B")
    $pendingColor = [System.Drawing.ColorTranslator]::FromHtml("#A45F00")

    $form = New-Object System.Windows.Forms.Form
    $form.Text = "PrintEdge"
    $form.Size = New-Object System.Drawing.Size(760, 690)
    $form.MinimumSize = New-Object System.Drawing.Size(620, 560)
    $form.StartPosition = "CenterScreen"
    $form.BackColor = $backgroundColor
    $form.Font = New-Object System.Drawing.Font("Segoe UI", 10)

    $headerPanel = New-Object System.Windows.Forms.Panel
    $headerPanel.Dock = "Top"
    $headerPanel.Height = 92
    $headerPanel.BackColor = $cardColor
    $headerPanel.Padding = New-Object System.Windows.Forms.Padding(22, 16, 22, 12)

    $titleLabel = New-Object System.Windows.Forms.Label
    $titleLabel.Text = "Active Print Queues"
    $titleLabel.AutoSize = $true
    $titleLabel.Font = New-Object System.Drawing.Font("Segoe UI Semibold", 17)
    $titleLabel.ForeColor = $textColor
    $titleLabel.Location = New-Object System.Drawing.Point(22, 16)

    $subtitleLabel = New-Object System.Windows.Forms.Label
    $subtitleLabel.AutoSize = $true
    $subtitleLabel.Font = New-Object System.Drawing.Font("Segoe UI", 10)
    $subtitleLabel.ForeColor = $mutedColor
    $subtitleLabel.Location = New-Object System.Drawing.Point(24, 54)

    if ($profile) {
        $profileName = Get-PropertyValue -InputObject $profile -Name "name" -Default "Unnamed profile"
        $profileCidr = Get-PropertyValue -InputObject $profile -Name "cidr" -Default "Unknown subnet"
        $subtitleLabel.Text = "{0}  |  {1}  |  {2}" -f $profileName, $profileCidr, $currentIPv4Address
    }
    else {
        $subtitleLabel.Text = "No configured profile matched {0}" -f $currentIPv4Address
    }

    $headerPanel.Controls.Add($titleLabel)
    $headerPanel.Controls.Add($subtitleLabel)

    $contentPanel = New-Object System.Windows.Forms.FlowLayoutPanel
    $contentPanel.Dock = "Fill"
    $contentPanel.AutoScroll = $true
    $contentPanel.FlowDirection = "TopDown"
    $contentPanel.WrapContents = $false
    $contentPanel.Padding = New-Object System.Windows.Forms.Padding(22, 18, 22, 18)
    $contentPanel.BackColor = $backgroundColor

    if ($null -eq $profile) {
        $emptyLabel = New-Object System.Windows.Forms.Label
        $emptyLabel.Text = "No active PrintEdge printers for this subnet."
        $emptyLabel.AutoSize = $true
        $emptyLabel.Font = New-Object System.Drawing.Font("Segoe UI", 12)
        $emptyLabel.ForeColor = $mutedColor
        $emptyLabel.Margin = New-Object System.Windows.Forms.Padding(0, 20, 0, 0)
        $contentPanel.Controls.Add($emptyLabel)
    }
    else {
        $desiredQueues = @(Get-DesiredPrintQueues -Profile $profile)
        foreach ($printerGroup in ($desiredQueues | Group-Object -Property PrinterId)) {
            $groupQueues = @($printerGroup.Group)
            $firstQueue = $groupQueues[0]

            $card = New-Object System.Windows.Forms.Panel
            $card.Width = 690
            $card.Height = 78 + (36 * $groupQueues.Count)
            $card.BackColor = $cardColor
            $card.Margin = New-Object System.Windows.Forms.Padding(0, 0, 0, 14)
            $card.Padding = New-Object System.Windows.Forms.Padding(18)
            $card.Add_Paint({
                param($sender, $eventArgs)
                $borderPen = New-Object System.Drawing.Pen([System.Drawing.ColorTranslator]::FromHtml("#DDE4EE"), 1)
                try {
                    $rect = $sender.ClientRectangle
                    $rect.Width = $rect.Width - 1
                    $rect.Height = $rect.Height - 1
                    $eventArgs.Graphics.DrawRectangle($borderPen, $rect)
                }
                finally {
                    $borderPen.Dispose()
                }
            })

            $printerLabel = New-Object System.Windows.Forms.Label
            $printerLabel.Text = $firstQueue.PrinterDisplayName
            $printerLabel.Font = New-Object System.Drawing.Font("Segoe UI Semibold", 12)
            $printerLabel.ForeColor = $textColor
            $printerLabel.Location = New-Object System.Drawing.Point(18, 14)
            $printerLabel.Size = New-Object System.Drawing.Size(430, 28)

            $addressLabel = New-Object System.Windows.Forms.Label
            $addressLabel.Text = $firstQueue.HostAddress
            $addressLabel.Font = New-Object System.Drawing.Font("Segoe UI", 9)
            $addressLabel.ForeColor = $mutedColor
            $addressLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleRight
            $addressLabel.Location = New-Object System.Drawing.Point(470, 14)
            $addressLabel.Size = New-Object System.Drawing.Size(190, 24)

            $locationLabel = New-Object System.Windows.Forms.Label
            $locationLabel.Text = $firstQueue.Location
            $locationLabel.Font = New-Object System.Drawing.Font("Segoe UI", 9)
            $locationLabel.ForeColor = $mutedColor
            $locationLabel.Location = New-Object System.Drawing.Point(18, 42)
            $locationLabel.Size = New-Object System.Drawing.Size(430, 22)

            $card.Controls.Add($printerLabel)
            $card.Controls.Add($addressLabel)
            $card.Controls.Add($locationLabel)

            $queueTop = 72
            foreach ($queue in $groupQueues) {
                $statusText = "Pending"
                $statusColor = $pendingColor
                try {
                    if (Get-Printer -Name $queue.Name -ErrorAction SilentlyContinue) {
                        $statusText = "Installed"
                        $statusColor = $successColor
                    }
                }
                catch {
                    $statusText = "Unknown"
                    $statusColor = $mutedColor
                }

                $queueNameLabel = New-Object System.Windows.Forms.Label
                $queueNameLabel.Text = $queue.Name
                $queueNameLabel.Font = New-Object System.Drawing.Font("Segoe UI", 10)
                $queueNameLabel.ForeColor = $textColor
                $queueNameLabel.Location = New-Object System.Drawing.Point(34, $queueTop)
                $queueNameLabel.Size = New-Object System.Drawing.Size(460, 24)

                $statusLabel = New-Object System.Windows.Forms.Label
                $statusLabel.Text = $statusText
                $statusLabel.Font = New-Object System.Drawing.Font("Segoe UI Semibold", 9)
                $statusLabel.ForeColor = $statusColor
                $statusLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleRight
                $statusLabel.Location = New-Object System.Drawing.Point(520, $queueTop)
                $statusLabel.Size = New-Object System.Drawing.Size(140, 24)

                $dot = New-Object System.Windows.Forms.Panel
                $dot.BackColor = $accentColor
                $dot.Location = New-Object System.Drawing.Point(18, ($queueTop + 8))
                $dot.Size = New-Object System.Drawing.Size(7, 7)

                $card.Controls.Add($dot)
                $card.Controls.Add($queueNameLabel)
                $card.Controls.Add($statusLabel)
                $queueTop += 34
            }

            $contentPanel.Controls.Add($card)
        }
    }

    $diagnosticPanel = New-Object System.Windows.Forms.Panel
    $diagnosticPanel.Dock = "Bottom"
    $diagnosticPanel.Height = 176
    $diagnosticPanel.BackColor = [System.Drawing.ColorTranslator]::FromHtml("#101827")
    $diagnosticPanel.Padding = New-Object System.Windows.Forms.Padding(18, 10, 18, 14)

    $diagnosticHeader = New-Object System.Windows.Forms.Label
    $diagnosticHeader.Text = "Console"
    $diagnosticHeader.AutoSize = $true
    $diagnosticHeader.Font = New-Object System.Drawing.Font("Segoe UI Semibold", 9)
    $diagnosticHeader.ForeColor = [System.Drawing.ColorTranslator]::FromHtml("#B7C4D6")
    $diagnosticHeader.Location = New-Object System.Drawing.Point(18, 8)

    $diagnosticTextBox = New-Object System.Windows.Forms.TextBox
    $diagnosticTextBox.Multiline = $true
    $diagnosticTextBox.ReadOnly = $true
    $diagnosticTextBox.ScrollBars = [System.Windows.Forms.ScrollBars]::Vertical
    $diagnosticTextBox.BorderStyle = [System.Windows.Forms.BorderStyle]::None
    $diagnosticTextBox.BackColor = [System.Drawing.ColorTranslator]::FromHtml("#101827")
    $diagnosticTextBox.ForeColor = [System.Drawing.ColorTranslator]::FromHtml("#E8EEF7")
    $diagnosticTextBox.Font = New-Object System.Drawing.Font("Consolas", 9)
    $diagnosticTextBox.Location = New-Object System.Drawing.Point(18, 30)
    $diagnosticTextBox.Size = New-Object System.Drawing.Size(704, 126)
    $diagnosticTextBox.Anchor = [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Right -bor [System.Windows.Forms.AnchorStyles]::Bottom
    $diagnosticTextBox.Text = Get-PrintEdgeDiagnosticText -Profile $profile -CurrentIPv4Address $currentIPv4Address
    $diagnosticTextBox.SelectionStart = $diagnosticTextBox.TextLength
    $diagnosticTextBox.ScrollToCaret()

    $diagnosticTimer = New-Object System.Windows.Forms.Timer
    $diagnosticTimer.Interval = 3000
    $diagnosticTimer.Add_Tick({
        $currentDialogIPv4Address = Get-ActiveIPv4Address
        $currentDialogProfile = $null
        if (-not [string]::IsNullOrWhiteSpace($currentDialogIPv4Address)) {
            $currentDialogProfile = Select-SubnetProfile -Config $Config -IPAddress $currentDialogIPv4Address
        }

        $diagnosticTextBox.Text = Get-PrintEdgeDiagnosticText -Profile $currentDialogProfile -CurrentIPv4Address $currentDialogIPv4Address
        $diagnosticTextBox.SelectionStart = $diagnosticTextBox.TextLength
        $diagnosticTextBox.ScrollToCaret()
    })
    $diagnosticTimer.Start()

    $form.Add_FormClosed({
        $diagnosticTimer.Stop()
        $diagnosticTimer.Dispose()
    })

    $diagnosticPanel.Controls.Add($diagnosticHeader)
    $diagnosticPanel.Controls.Add($diagnosticTextBox)

    $buttonPanel = New-Object System.Windows.Forms.Panel
    $buttonPanel.Dock = "Bottom"
    $buttonPanel.Height = 58
    $buttonPanel.BackColor = $cardColor

    $okButton = New-Object System.Windows.Forms.Button
    $okButton.Text = "Close"
    $okButton.Width = 96
    $okButton.Height = 32
    $okButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $okButton.BackColor = $accentColor
    $okButton.ForeColor = [System.Drawing.Color]::White
    $okButton.FlatAppearance.BorderSize = 0
    $okButton.Location = New-Object System.Drawing.Point(638, 13)
    $okButton.Anchor = [System.Windows.Forms.AnchorStyles]::Right -bor [System.Windows.Forms.AnchorStyles]::Top
    $okButton.Add_Click({ $form.Close() })

    $buttonPanel.Controls.Add($okButton)
    $form.Controls.Add($contentPanel)
    $form.Controls.Add($diagnosticPanel)
    $form.Controls.Add($buttonPanel)
    $form.Controls.Add($headerPanel)
    $form.ShowDialog() | Out-Null
}

function Show-ActiveSubnetDialog {
    $currentIPv4Address = Get-ActiveIPv4Address
    $profile = $null
    if (-not [string]::IsNullOrWhiteSpace($currentIPv4Address)) {
        $profile = Select-SubnetProfile -Config $script:Config -IPAddress $currentIPv4Address
    }

    $profileName = "No matching profile"
    if ($profile) {
        $profileName = Get-PropertyValue -InputObject $profile -Name "name" -Default "Unnamed profile"
    }

    [System.Windows.Forms.MessageBox]::Show(
        ("IPv4 Address: {0}`r`nProfile: {1}" -f $currentIPv4Address, $profileName),
        "PrintEdge Active Network",
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Information
    ) | Out-Null
}

function Open-PrintEdgeConfig {
    param(
        [Parameter(Mandatory)]
        [string]$Path
    )

    $resolvedPath = Expand-PrintEdgePath -Path $Path
    if (Test-Path -Path $resolvedPath) {
        Start-Process -FilePath "notepad.exe" -ArgumentList @($resolvedPath)
        return
    }

    [System.Windows.Forms.MessageBox]::Show(
        ("Config file not found: {0}" -f $resolvedPath),
        "PrintEdge Settings",
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Warning
    ) | Out-Null
}

function Set-NotifyIconText {
    param(
        [Parameter(Mandatory)]
        [System.Windows.Forms.NotifyIcon]$NotifyIcon,
        [Parameter(Mandatory)]
        [string]$Text
    )

    if ($Text.Length -gt 63) {
        $Text = $Text.Substring(0, 60) + "..."
    }

    $NotifyIcon.Text = $Text
}

function Update-PrintEdgeTrayStatus {
    param(
        [Parameter(Mandatory)]
        [object]$Config,
        [Parameter(Mandatory)]
        [System.Windows.Forms.NotifyIcon]$NotifyIcon,
        [Parameter(Mandatory)]
        [System.Windows.Forms.MenuItem]$SubnetMenuItem,
        [Parameter(Mandatory)]
        [System.Windows.Forms.MenuItem]$ProfileMenuItem
    )

    $currentIPv4Address = Get-ActiveIPv4Address
    if ([string]::IsNullOrWhiteSpace($currentIPv4Address)) {
        $SubnetMenuItem.Text = "Subnet: Not connected"
        $ProfileMenuItem.Text = "Active queues: 0"
        Set-NotifyIconText -NotifyIcon $NotifyIcon -Text "PrintEdge - not connected"
        return
    }

    $profile = Select-SubnetProfile -Config $Config -IPAddress $currentIPv4Address
    if ($null -eq $profile) {
        $SubnetMenuItem.Text = "Subnet: No configured match"
        $ProfileMenuItem.Text = "IP: {0}" -f $currentIPv4Address
        Set-NotifyIconText -NotifyIcon $NotifyIcon -Text "PrintEdge - no subnet match"
        return
    }

    $profileName = Get-PropertyValue -InputObject $profile -Name "name" -Default "Unnamed profile"
    $profileCidr = Get-PropertyValue -InputObject $profile -Name "cidr" -Default $currentIPv4Address
    $queueCount = @(Get-DesiredPrintQueues -Profile $profile).Count

    $SubnetMenuItem.Text = "Subnet: {0}" -f $profileCidr
    $ProfileMenuItem.Text = "{0} - {1} active queue(s)" -f $profileName, $queueCount
    Set-NotifyIconText -NotifyIcon $NotifyIcon -Text ("PrintEdge - {0}" -f $profileName)
}

function Start-PrintEdgeTray {
    param(
        [Parameter(Mandatory)]
        [object]$Config,
        [switch]$HideConsole,
        [switch]$DryRun
    )

    if ($HideConsole) {
        Hide-ConsoleWindow
    }

    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing
    [System.Windows.Forms.Application]::EnableVisualStyles()

    $notifyIcon = New-Object System.Windows.Forms.NotifyIcon
    $notifyIcon.Icon = New-PrintEdgeTrayIcon
    $notifyIcon.Text = "PrintEdge"
    $notifyIcon.Visible = $true
    $notifyIcon.add_DoubleClick({ Show-PrintersDialog -Config $script:Config })

    $contextMenu = New-Object System.Windows.Forms.ContextMenu

    $menuSubnet = New-Object System.Windows.Forms.MenuItem("Subnet: Detecting...")
    $menuSubnet.Enabled = $false

    $menuProfile = New-Object System.Windows.Forms.MenuItem("Active queues: Detecting...")
    $menuProfile.Enabled = $false

    $menuSyncNow = New-Object System.Windows.Forms.MenuItem("Sync Now")
    $menuSyncNow.add_Click({
        try {
            Sync-CurrentSubnet -Force -DryRun:$DryRun
            Update-PrintEdgeTrayStatus -Config $script:Config -NotifyIcon $notifyIcon -SubnetMenuItem $menuSubnet -ProfileMenuItem $menuProfile
            [System.Windows.Forms.MessageBox]::Show(
                "Printer sync completed or elevated sync was requested.",
                "PrintEdge",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Information
            ) | Out-Null
        }
        catch {
            Write-PrintEdgeLog -Level "ERROR" -Message $_.Exception.Message
            [System.Windows.Forms.MessageBox]::Show(
                $_.Exception.Message,
                "PrintEdge Sync Failed",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Error
            ) | Out-Null
        }
    })

    $menuExit = New-Object System.Windows.Forms.MenuItem("Quit")
    $menuExit.add_Click({
        $icon = $notifyIcon.Icon
        $notifyIcon.Visible = $false
        $notifyIcon.Dispose()
        if ($icon) {
            $icon.Dispose()
        }
        [System.Windows.Forms.Application]::Exit()
    })

    $contextMenu.MenuItems.Add($menuSubnet) | Out-Null
    $contextMenu.MenuItems.Add($menuProfile) | Out-Null
    $contextMenu.MenuItems.Add("-") | Out-Null
    $contextMenu.MenuItems.Add($menuSyncNow) | Out-Null
    $contextMenu.MenuItems.Add($menuExit) | Out-Null
    $notifyIcon.ContextMenu = $contextMenu
    $notifyIcon.add_MouseUp({
        if ($_.Button -eq [System.Windows.Forms.MouseButtons]::Right) {
            Update-PrintEdgeTrayStatus -Config $script:Config -NotifyIcon $notifyIcon -SubnetMenuItem $menuSubnet -ProfileMenuItem $menuProfile
        }
    })

    $pollIntervalSeconds = [int](Get-Setting -Name "pollIntervalSeconds" -Default 30)
    if ($pollIntervalSeconds -lt 10) {
        $pollIntervalSeconds = 10
    }

    $timer = New-Object System.Windows.Forms.Timer
    $timer.Interval = $pollIntervalSeconds * 1000
    $timer.Add_Tick({
        try {
            Sync-CurrentSubnet -DryRun:$DryRun
            Update-PrintEdgeTrayStatus -Config $script:Config -NotifyIcon $notifyIcon -SubnetMenuItem $menuSubnet -ProfileMenuItem $menuProfile
        }
        catch {
            Write-PrintEdgeLog -Level "ERROR" -Message $_.Exception.Message
        }
    })
    $timer.Start()

    try {
        Sync-CurrentSubnet -Force -DryRun:$DryRun
        Update-PrintEdgeTrayStatus -Config $script:Config -NotifyIcon $notifyIcon -SubnetMenuItem $menuSubnet -ProfileMenuItem $menuProfile
    }
    catch {
        Write-PrintEdgeLog -Level "ERROR" -Message $_.Exception.Message
    }

    [System.Windows.Forms.Application]::Run()
}

function Main {
    $config = Import-PrintEdgeConfig -Path $ConfigPath
    Write-PrintEdgeLog -Message ("Loaded configuration from '{0}'." -f (Expand-PrintEdgePath -Path $ConfigPath))

    if ($ValidateConfig) {
        Write-Output "PrintEdge configuration is valid."
        return
    }

    $dryRun = $WhatIfMode -or (ConvertTo-Bool -Value (Get-Setting -Name "dryRun" -Default $false) -Default $false)

    if ($Once -or $NoTray) {
        Sync-CurrentSubnet -Force -DryRun:$dryRun
        return
    }

    Start-PrintEdgeTray -Config $config -HideConsole:(-not $ShowConsole) -DryRun:$dryRun
}

Main
