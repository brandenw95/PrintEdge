[CmdletBinding()]
param(
    [string]$InstallPath = "",
    [switch]$Uninstall,
    [switch]$WhatIfMode
)

Set-StrictMode -Version 2.0
$ErrorActionPreference = "Stop"

$appName = "PrintEdge"
$taskFolderPath = "\PrintEdge"
$syncTaskName = "Sync"
$trayTaskName = "Tray"
$sourcePath = $PSScriptRoot

if ([string]::IsNullOrWhiteSpace($InstallPath)) {
    $InstallPath = Join-Path -Path $env:ProgramFiles -ChildPath $appName
}

function Write-InstallLog {
    param(
        [Parameter(Mandatory)]
        [string]$Message,
        [ValidateSet("INFO", "WARN", "ERROR")]
        [string]$Level = "INFO"
    )

    $line = "{0} [{1}] {2}" -f (Get-Date -Format "yyyy-MM-dd HH:mm:ss"), $Level, $Message
    Write-Host $line
}

function Test-IsAdministrator {
    $identity = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object Security.Principal.WindowsPrincipal($identity)
    return $principal.IsInRole([Security.Principal.WindowsBuiltinRole]::Administrator)
}

function Get-PowerShellPath {
    $powerShellPath = Join-Path -Path $env:WINDIR -ChildPath "System32\WindowsPowerShell\v1.0\powershell.exe"
    if (Test-Path -Path $powerShellPath) {
        return $powerShellPath
    }

    return "powershell.exe"
}

function Join-PowerShellArguments {
    param(
        [string[]]$Arguments
    )

    return (($Arguments | ForEach-Object {
        if ($_ -match "\s") {
            '"{0}"' -f ($_ -replace '"', '\"')
        }
        else {
            $_
        }
    }) -join " ")
}

function Ensure-TaskFolder {
    $service = New-Object -ComObject "Schedule.Service"
    $service.Connect()
    $rootFolder = $service.GetFolder("\")
    $folderName = $taskFolderPath.Trim("\")

    try {
        $rootFolder.GetFolder($folderName) | Out-Null
    }
    catch {
        $rootFolder.CreateFolder($folderName) | Out-Null
    }
}

function Grant-TaskRunAccess {
    param(
        [Parameter(Mandatory)]
        [string]$TaskName
    )

    $service = New-Object -ComObject "Schedule.Service"
    $service.Connect()
    $folder = $service.GetFolder($taskFolderPath)
    $task = $folder.GetTask($TaskName)
    $sddl = [string]$task.GetSecurityDescriptor(0)
    $authenticatedUsersExecuteAce = "(A;;0x1200a9;;;AU)"

    if ($sddl -like "*$authenticatedUsersExecuteAce*") {
        return
    }

    $daclIndex = $sddl.IndexOf("D:")
    if ($daclIndex -lt 0) {
        $updatedSddl = "{0}D:{1}" -f $sddl, $authenticatedUsersExecuteAce
    }
    else {
        $saclIndex = $sddl.IndexOf("S:", $daclIndex)
        if ($saclIndex -gt -1) {
            $updatedSddl = $sddl.Insert($saclIndex, $authenticatedUsersExecuteAce)
        }
        else {
            $updatedSddl = $sddl + $authenticatedUsersExecuteAce
        }
    }

    $task.SetSecurityDescriptor($updatedSddl, 0)
}

function Install-AppFiles {
    if (-not (Test-Path -Path $InstallPath)) {
        New-Item -Path $InstallPath -ItemType Directory -Force | Out-Null
    }

    foreach ($fileName in @("PrintEdge.ps1", "PrintEdge.config.json", "PrintEdge.ico")) {
        $sourceFile = Join-Path -Path $sourcePath -ChildPath $fileName
        if (-not (Test-Path -Path $sourceFile)) {
            throw "Required source file not found: $sourceFile"
        }

        Copy-Item -Path $sourceFile -Destination (Join-Path -Path $InstallPath -ChildPath $fileName) -Force
    }

    $acl = Get-Acl -Path $InstallPath
    $acl.SetAccessRuleProtection($true, $true)
    $usersRule = New-Object System.Security.AccessControl.FileSystemAccessRule(
        "BUILTIN\Users",
        "ReadAndExecute",
        "ContainerInherit,ObjectInherit",
        "None",
        "Allow"
    )
    $acl.SetAccessRule($usersRule)
    Set-Acl -Path $InstallPath -AclObject $acl

    Write-InstallLog -Message ("Copied PrintEdge files to '{0}'." -f $InstallPath)
}

function Install-ScheduledTasks {
    $printEdgeScript = Join-Path -Path $InstallPath -ChildPath "PrintEdge.ps1"
    $printEdgeConfig = Join-Path -Path $InstallPath -ChildPath "PrintEdge.config.json"
    $powerShellPath = Get-PowerShellPath
    $taskPathForCmdlets = "{0}\" -f $taskFolderPath

    Ensure-TaskFolder

    $syncArguments = Join-PowerShellArguments -Arguments @(
        "-NoProfile",
        "-ExecutionPolicy",
        "Bypass",
        "-WindowStyle",
        "Hidden",
        "-File",
        $printEdgeScript,
        "-ConfigPath",
        $printEdgeConfig,
        "-Once"
    )

    if ($WhatIfMode) {
        $syncArguments = "{0} -WhatIfMode" -f $syncArguments
    }

    $syncAction = New-ScheduledTaskAction -Execute $powerShellPath -Argument $syncArguments
    $syncTriggers = @(
        (New-ScheduledTaskTrigger -AtStartup),
        (New-ScheduledTaskTrigger -AtLogOn)
    )
    $syncPrincipal = New-ScheduledTaskPrincipal -UserId "SYSTEM" -LogonType ServiceAccount -RunLevel Highest
    $syncSettings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -MultipleInstances IgnoreNew -StartWhenAvailable
    $syncTask = New-ScheduledTask -Action $syncAction -Trigger $syncTriggers -Principal $syncPrincipal -Settings $syncSettings -Description "Runs PrintEdge printer synchronization with SYSTEM privileges."
    Register-ScheduledTask -TaskPath $taskPathForCmdlets -TaskName $syncTaskName -InputObject $syncTask -Force | Out-Null
    Grant-TaskRunAccess -TaskName $syncTaskName

    $trayArguments = Join-PowerShellArguments -Arguments @(
        "-NoProfile",
        "-ExecutionPolicy",
        "Bypass",
        "-WindowStyle",
        "Hidden",
        "-File",
        $printEdgeScript,
        "-ConfigPath",
        $printEdgeConfig
    )

    $trayAction = New-ScheduledTaskAction -Execute $powerShellPath -Argument $trayArguments
    $trayTrigger = New-ScheduledTaskTrigger -AtLogOn
    $trayPrincipal = New-ScheduledTaskPrincipal -GroupId "S-1-5-32-545" -RunLevel Limited
    $traySettings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -MultipleInstances IgnoreNew
    $trayTask = New-ScheduledTask -Action $trayAction -Trigger $trayTrigger -Principal $trayPrincipal -Settings $traySettings -Description "Starts the PrintEdge tray UI for signed-in standard users."
    Register-ScheduledTask -TaskPath $taskPathForCmdlets -TaskName $trayTaskName -InputObject $trayTask -Force | Out-Null

    Write-InstallLog -Message ("Created scheduled tasks '{0}\{1}' and '{0}\{2}'." -f $taskFolderPath, $syncTaskName, $trayTaskName)
}

function Install-Shortcuts {
    $programsPath = [Environment]::GetFolderPath("CommonPrograms")
    $shortcutFolder = Join-Path -Path $programsPath -ChildPath $appName
    if (-not (Test-Path -Path $shortcutFolder)) {
        New-Item -Path $shortcutFolder -ItemType Directory -Force | Out-Null
    }

    $printEdgeScript = Join-Path -Path $InstallPath -ChildPath "PrintEdge.ps1"
    $printEdgeConfig = Join-Path -Path $InstallPath -ChildPath "PrintEdge.config.json"
    $iconPath = Join-Path -Path $InstallPath -ChildPath "PrintEdge.ico"

    $powerShellPath = Get-PowerShellPath
    $shell = New-Object -ComObject "WScript.Shell"

    $trayShortcut = $shell.CreateShortcut((Join-Path -Path $shortcutFolder -ChildPath "PrintEdge.lnk"))
    $trayShortcut.TargetPath = $powerShellPath
    $trayShortcut.Arguments = Join-PowerShellArguments -Arguments @(
        "-NoProfile",
        "-ExecutionPolicy",
        "Bypass",
        "-WindowStyle",
        "Hidden",
        "-File",
        $printEdgeScript,
        "-ConfigPath",
        $printEdgeConfig
    )
    $trayShortcut.WorkingDirectory = $InstallPath
    $trayShortcut.Description = "Start the PrintEdge tray application."
    $trayShortcut.IconLocation = $iconPath
    $trayShortcut.Save()

    $syncShortcut = $shell.CreateShortcut((Join-Path -Path $shortcutFolder -ChildPath "PrintEdge Sync Now.lnk"))
    $syncShortcut.TargetPath = Join-Path -Path $env:WINDIR -ChildPath "System32\schtasks.exe"
    $syncShortcut.Arguments = '/Run /TN "\PrintEdge\Sync"'
    $syncShortcut.WorkingDirectory = $InstallPath
    $syncShortcut.Description = "Run PrintEdge printer sync with SYSTEM privileges."
    $syncShortcut.IconLocation = $iconPath
    $syncShortcut.Save()

    Write-InstallLog -Message ("Created Start Menu shortcuts in '{0}'." -f $shortcutFolder)
}

function Start-PrintEdge {
    $syncTaskPath = "{0}\{1}" -f $taskFolderPath, $syncTaskName
    $process = Start-Process -FilePath "schtasks.exe" -ArgumentList @("/Run", "/TN", $syncTaskPath) -WindowStyle Hidden -Wait -PassThru
    if ($process.ExitCode -eq 0) {
        Write-InstallLog -Message ("Started elevated sync task '{0}'." -f $syncTaskPath)
    }
    else {
        Write-InstallLog -Level "WARN" -Message ("Unable to start elevated sync task '{0}'. schtasks exit code: {1}" -f $syncTaskPath, $process.ExitCode)
    }
}

function Start-PrintEdgeTray {
    $trayTaskPath = "{0}\{1}" -f $taskFolderPath, $trayTaskName
    $process = Start-Process -FilePath "schtasks.exe" -ArgumentList @("/Run", "/TN", $trayTaskPath) -WindowStyle Hidden -Wait -PassThru
    if ($process.ExitCode -eq 0) {
        Write-InstallLog -Message ("Started tray task '{0}'." -f $trayTaskPath)
    }
    else {
        Write-InstallLog -Level "WARN" -Message ("Unable to start tray task '{0}'. It will start at the next user logon. schtasks exit code: {1}" -f $trayTaskPath, $process.ExitCode)
    }
}

function Remove-Shortcuts {
    $programsPath = [Environment]::GetFolderPath("CommonPrograms")
    $shortcutFolder = Join-Path -Path $programsPath -ChildPath $appName
    if (Test-Path -Path $shortcutFolder) {
        Remove-Item -Path $shortcutFolder -Recurse -Force
        Write-InstallLog -Message ("Removed Start Menu shortcuts from '{0}'." -f $shortcutFolder)
    }
}

function Remove-ScheduledTasks {
    $taskPathForCmdlets = "{0}\" -f $taskFolderPath
    foreach ($taskName in @($syncTaskName, $trayTaskName)) {
        try {
            Unregister-ScheduledTask -TaskPath $taskPathForCmdlets -TaskName $taskName -Confirm:$false -ErrorAction Stop
            Write-InstallLog -Message ("Removed scheduled task '{0}\{1}'." -f $taskFolderPath, $taskName)
        }
        catch {
            Write-InstallLog -Level "WARN" -Message ("Scheduled task '{0}\{1}' was not removed: {2}" -f $taskFolderPath, $taskName, $_.Exception.Message)
        }
    }
}

function Remove-AppFiles {
    if (Test-Path -Path $InstallPath) {
        Remove-Item -Path $InstallPath -Recurse -Force
        Write-InstallLog -Message ("Removed '{0}'." -f $InstallPath)
    }
}

if ($env:PROCESSOR_ARCHITEW6432) {
    $sysnativePowerShell = Join-Path -Path $env:WINDIR -ChildPath "Sysnative\WindowsPowerShell\v1.0\powershell.exe"
    if (Test-Path -Path $sysnativePowerShell) {
        $arguments = @(
            "-NoProfile",
            "-ExecutionPolicy",
            "Bypass",
            "-File",
            $PSCommandPath,
            "-InstallPath",
            $InstallPath
        )

        if ($Uninstall) {
            $arguments += "-Uninstall"
        }

        if ($WhatIfMode) {
            $arguments += "-WhatIfMode"
        }

        $relaunch = Start-Process -FilePath $sysnativePowerShell -ArgumentList $arguments -Wait -PassThru
        exit $relaunch.ExitCode
    }
}

if (-not (Test-IsAdministrator)) {
    throw "Install.ps1 must run elevated. In Intune, deploy it as a Win32 app with install behavior set to System."
}

if ($Uninstall) {
    Remove-Shortcuts
    Remove-ScheduledTasks
    Remove-AppFiles
    exit 0
}

Install-AppFiles
Install-ScheduledTasks
Install-Shortcuts
Start-PrintEdge
Start-PrintEdgeTray
