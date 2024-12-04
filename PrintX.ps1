# Imports
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
function Show-PrintersDialog {

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

function Open-Settings {

    # Show(Message Text, Title Text, OK Button, Information Icon)
    [System.Windows.Forms.MessageBox]::Show("Settings functionality is a placeholder.",
                                            "Settings", 
                                            [System.Windows.Forms.MessageBoxButtons]::OK, 
                                            [System.Windows.Forms.MessageBoxIcon]::Information
                                            )
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

    $menuViewPrinters = New-Object System.Windows.Forms.MenuItem("View Printers")
    $menuSettings = New-Object System.Windows.Forms.MenuItem("Settings")
    $menuExit = New-Object System.Windows.Forms.MenuItem("Exit")

    $menuViewPrinters.add_Click({ Show-PrintersDialog })
    $menuSettings.add_Click({ Open-Settings })
    $menuExit.add_Click({
        $notifyIcon.Visible = $false
        [System.Windows.Forms.Application]::Exit()
    })

    $contextMenu.MenuItems.Add($menuViewPrinters) | Out-Null
    $contextMenu.MenuItems.Add($menuSettings) | Out-Null
    $contextMenu.MenuItems.Add($menuExit) | Out-Null

    $notifyIcon.ContextMenu = $contextMenu

    [System.Windows.Forms.Application]::Run()
}
Main