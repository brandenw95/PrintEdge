# Imports
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
function ShowPrintersDialog {

    $printers = @(
        @{ Name = "Printer A"; Status = "Online" },
        @{ Name = "Printer B"; Status = "Offline" },
        @{ Name = "Printer C"; Status = "Online" }
    )

    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Printers"
    $form.Size = New-Object System.Drawing.Size(300, 200)
    $form.StartPosition = "CenterScreen"

    $listBox = New-Object System.Windows.Forms.ListBox
    $listBox.Dock = "Fill"
    $listBox.Items.AddRange(($printers | ForEach-Object { "$($_.Name) - $($_.Status)" }) -as [string[]]) | Out-Null

    $form.Controls.Add($listBox)

    $okButton = New-Object System.Windows.Forms.Button
    $okButton.Text = "OK"
    $okButton.Dock = "Bottom"
    $okButton.Add_Click({ $form.Close() })
    $form.Controls.Add($okButton)

    $form.ShowDialog()
}

function OpenSettings {

    # Show(Message Text, Title Text, OK Button, Information Icon)
    [System.Windows.Forms.MessageBox]::Show("Settings functionality is a placeholder.",
                                            "Settings", 
                                            [System.Windows.Forms.MessageBoxButtons]::OK, 
                                            [System.Windows.Forms.MessageBoxIcon]::Information
                                            )
}

function InstallPrinters {
    param (
        [Parameter(Mandatory)]
        [hashtable]$Printers,
        [Parameter(Mandatory)]
        [string]$BaseIP
    )

    # EXAMPLE FUNCTION INPUT
    # Example usage:
    # $Printers = @{
    #     "Brother HL-L2320D series" = @{name = "Brother HL-L2320D series"; inf="drivers\Brother-HL-L2320D\32_64\BROHL13A.INF"; ip="192.168.1.110"}
    #     "Brother HL-L2370DW series" = @{name = "Brother HL-L2370DW series"; inf="drivers\Brother-HL-L2370DW\gdi\BROHL17A.INF"; ip="192.168.1.100"}
    # }

    foreach ($Printer in $Printers.GetEnumerator()) {
        $PrinterName = $Printer.Value.name
        $DriverPath = $Printer.Value.inf
        $CurrentIP = $BaseIP

        # START - DRIVER INSTALL
        Write-Host "Installing driver for $PrinterName from $DriverPath..."
        pnputil /add-driver $DriverPath /install | Out-Null

        # START - IP PORT CONFIG CHECK
        $PortName = "IP_$CurrentIP"
        $Counter = 1
        while (Get-PrinterPort -Name $PortName -ErrorAction SilentlyContinue) {
            $PortName = "IP_$BaseIP_$Counter"
            $Counter++
        }

        Write-Host "Creating printer port $PortName for IP $CurrentIP..."
        Add-PrinterPort -Name $PortName -PrinterHostAddress $CurrentIP

        # START - ADD PRINTER TO LIST
        $PrinterFinalName = $PrinterName
        $Suffix = 1
        while (Get-Printer -Name $PrinterFinalName -ErrorAction SilentlyContinue) {
            $PrinterFinalName = "$PrinterName_$Suffix"
            $Suffix++
        }

        Write-Host "Adding printer $PrinterFinalName with port $PortName..."
        Add-Printer -Name $PrinterFinalName -DriverName $PrinterName -PortName $PortName

        $BaseIP = ([IPV4Address]::Parse($BaseIP).GetAddressBytes() | ForEach-Object{ [int]$_ } | ForEach-Object { $_.ToString() }) -join '.'
    }

    Write-Host "All printers installed successfully!"
}

function GetSubnet{
    # Get the current active network adapter with IPv4 addresses
    $currentAdapter = Get-NetIPAddress -AddressFamily IPv4 |
        Where-Object {
            $_.IPAddress -ne "127.0.0.1"
        } |
        Sort-Object -Property InterfaceIndex | Select-Object -First 1

    if (-not $currentAdapter) {
        Write-Error "No active IPv4 network adapters found."
        return $null
    }

    # Return the current IP address
    #Write-Output "$currentAdapter.IPAddress"
    return $currentAdapter.IPAddress
}


Function Main{

    $scriptDirectory = $PSScriptRoot
    Write-Output $PSScriptRoot
    $iconPath = Join-Path -Path $scriptDirectory -ChildPath "tray_icon.ico"
    
    if (-Not (Test-Path -Path $iconPath)) {
        Write-Warning "Custom icon 'tray_icon.ico' not found in script directory. Using default PowerShell icon."
        $iconPath = "$($PSHOME)\powershell.exe"
    }

    $notifyIcon = New-Object System.Windows.Forms.NotifyIcon
    $notifyIcon.Icon = [System.Drawing.Icon]::ExtractAssociatedIcon($iconPath)
    $notifyIcon.Text = "Print Management"
    $notifyIcon.Visible = $true

    $contextMenu = New-Object System.Windows.Forms.ContextMenu

    # View Printers
    $menuViewPrinters = New-Object System.Windows.Forms.MenuItem("View Printers")
    $menuViewPrinters.add_Click({ ShowPrintersDialog })

    # Get Subnet
    $GetSubnet = New-Object System.Windows.Forms.MenuItem("Get Subnet")
    $GetSubnet.add_Click({
        $subnet = GetSubnet
        [System.Windows.Forms.MessageBox]::Show("Subnet: $subnet", "Active Subnet", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        Write-Output "Active Subnet: $subnet"
    })

    # Install Network Printers
    $InstallPrintSetting = New-Object System.Windows.Forms.MenuItem("Install Printers")
    $InstallPrintSetting.add_Click({ InstallPrinters })

    # View Settings
    $menuSettings = New-Object System.Windows.Forms.MenuItem("Settings")
    $menuSettings.add_Click({ OpenSettings })
    
    # Exit
    $menuExit = New-Object System.Windows.Forms.MenuItem("Exit")
    $menuExit.add_Click({
        $notifyIcon.Visible = $false
        [System.Windows.Forms.Application]::Exit()
    })

    $contextMenu.MenuItems.Add($menuViewPrinters) | Out-Null
    $contextMenu.MenuItems.Add($menuSettings) | Out-Null
    $contextMenu.MenuItems.Add($InstallPrintSetting) | Out-Null
    $contextMenu.MenuItems.Add($GetSubnet) | Out-Null
    $contextMenu.MenuItems.Add($menuExit) | Out-Null

    $notifyIcon.ContextMenu = $contextMenu

    [System.Windows.Forms.Application]::Run()
}
Main