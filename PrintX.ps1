# Import required namespaces
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Get the directory of the script and set the icon path
$scriptDirectory = Split-Path -Parent $MyInvocation.MyCommand.Path
$iconPath = Join-Path -Path $scriptDirectory -ChildPath "trayicon.ico"

# Verify if the icon file exists
if (-Not (Test-Path -Path $iconPath)) {
    Write-Warning "Custom icon 'trayicon.ico' not found in script directory. Using default PowerShell icon."
    $iconPath = "$($PSHOME)\powershell.exe"
}

# Create the NotifyIcon
$notifyIcon = New-Object System.Windows.Forms.NotifyIcon
$notifyIcon.Icon = [System.Drawing.Icon]::ExtractAssociatedIcon($iconPath)
$notifyIcon.Text = "Print Management"
$notifyIcon.Visible = $true

# Create Context Menu
$contextMenu = New-Object System.Windows.Forms.ContextMenu

# Dummy printer data
$printers = @(
    @{ Name = "Printer A"; Status = "Online" },
    @{ Name = "Printer B"; Status = "Offline" },
    @{ Name = "Printer C"; Status = "Online" }
)

# Function: Show Printers in a dialog
function Show-PrintersDialog {
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Printers"
    $form.Size = New-Object System.Drawing.Size(300, 200)
    $form.StartPosition = "CenterScreen"

    $listBox = New-Object System.Windows.Forms.ListBox
    $listBox.Dock = "Fill"
    $listBox.Items.AddRange(($printers | ForEach-Object { "$($_.Name) - $($_.Status)" }) -as [string[]])


    $form.Controls.Add($listBox)

    # Add an OK button to close the form
    $okButton = New-Object System.Windows.Forms.Button
    $okButton.Text = "OK"
    $okButton.Dock = "Bottom"
    $okButton.Add_Click({ $form.Close() })
    $form.Controls.Add($okButton)

    # Show the dialog
    $form.ShowDialog()
}

# Function: Open Settings (Placeholder)
function Open-Settings {
    [System.Windows.Forms.MessageBox]::Show("Settings functionality is a placeholder.", "Settings", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
}

# Add Menu Items
$menuViewPrinters = New-Object System.Windows.Forms.MenuItem("View Printers")
$menuSettings = New-Object System.Windows.Forms.MenuItem("Settings")
$menuExit = New-Object System.Windows.Forms.MenuItem("Exit")

# Add event handlers for menu items
$menuViewPrinters.add_Click({ Show-PrintersDialog })
$menuSettings.add_Click({ Open-Settings })
$menuExit.add_Click({
    $notifyIcon.Visible = $false
    [System.Windows.Forms.Application]::Exit()
})

# Add menu items to context menu
$contextMenu.MenuItems.Add($menuViewPrinters)
$contextMenu.MenuItems.Add($menuSettings)
$contextMenu.MenuItems.Add($menuExit)

# Assign context menu to NotifyIcon
$notifyIcon.ContextMenu = $contextMenu

# Run the application to handle events
[System.Windows.Forms.Application]::Run()
