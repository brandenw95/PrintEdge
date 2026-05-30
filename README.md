# PrintEdge

PrintEdge is a Windows PowerShell tray application for internal printer mapping. It detects the active IPv4 subnet, selects a matching profile from `PrintEdge.config.json`, and adds or removes the print queues assigned to that subnet.

This project is intended as a lightweight internal Printix-style client. The included configuration uses dummy printers and the built-in `Microsoft IPP Class Driver` so the app can be validated without vendor driver packages.

## Features

- Subnet-to-printer and subnet-to-queue mapping through `PrintEdge.config.json`.
- CIDR matching, for example `192.168.1.0/24`.
- Generated system tray icon with active subnet status, manual sync, and quit.
- Double-click tray popup showing only the active subnet's printers and queues.
- Safe cleanup that removes only PrintEdge-managed queues by default.
- Intune-friendly SYSTEM sync task for direct TCP/IP queues with a standard-user tray UI.
- Optional Azure Blob Storage driver download for `.zip` or `.inf` driver packages.
- CLI validation and one-shot sync modes for deployment automation.
- Separate Intune installer script that creates elevated scheduled tasks.
- Log output to `%ProgramData%\PrintEdge\Logs\PrintEdge.log` by default.

## Files

- `PrintX.ps1` - main PrintEdge tray and sync program.
- `PrintEdge.config.json` - subnet, printer, queue, driver, and Azure settings.
- `Install.ps1` - Intune/Win32 installer that copies PrintEdge to Program Files and creates scheduled tasks.

## Requirements

- Windows 10/11 or Windows Server with the Print Management PowerShell cmdlets.
- Windows PowerShell 5.1.
- Direct local TCP/IP queue creation and system driver installation are Windows admin operations.
- Standard users can run the tray UI after Intune installs the SYSTEM sync task with `Install.ps1`.

## Validate

Run this from the project folder:

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\PrintX.ps1 -ValidateConfig
```

## Run

Start the tray app:

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\PrintX.ps1
```

Run a single sync without the tray:

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\PrintX.ps1 -Once
```

Preview actions without installing drivers or changing printers:

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\PrintX.ps1 -Once -WhatIfMode
```

## Deploy

For Intune, package the repository folder as a Win32 app and use system install behavior. Use `Install.ps1` as the Intune install command:

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\Install.ps1
```

The installer copies these files to `C:\Program Files\PrintEdge`:

- `PrintX.ps1`
- `PrintEdge.config.json`

It then creates:

- `\PrintEdge\Sync` - runs `PrintX.ps1 -Once` as `SYSTEM` with highest privileges at startup, at user logon, and when requested.
- `\PrintEdge\Tray` - starts the normal `PrintX.ps1` tray UI for signed-in standard users.
- Start Menu shortcuts for **PrintEdge** and **PrintEdge Sync Now**.

Standard users can use **PrintEdge Sync Now** to trigger the SYSTEM sync task without being local administrators. The files live under Program Files, so users can run the approved task but cannot change the script or config that the task executes.

Use this Intune uninstall command:

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\Install.ps1 -Uninstall
```

If you install manually instead of through Intune, copy the two runtime files to `C:\Program Files\PrintEdge` first:

Validate the dropped files:

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "C:\Program Files\PrintEdge\PrintX.ps1" -ValidateConfig
```

You can still run the tray app manually from that location:

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File "C:\Program Files\PrintEdge\PrintX.ps1"
```

## Config

Each subnet profile contains physical printers and logical print queues:

```json
{
  "name": "Head Office",
  "cidr": "192.168.1.0/24",
  "printers": [
    {
      "id": "hq-mfp-01",
      "hostAddress": "192.168.1.50",
      "defaultDriverName": "Microsoft IPP Class Driver"
    }
  ],
  "printQueues": [
    {
      "name": "HQ Color MFP",
      "printerId": "hq-mfp-01",
      "comment": "Default color queue"
    }
  ]
}
```

Without `connectionName`, PrintEdge treats the entry as a direct local TCP/IP queue. Windows requires elevation to create local TCP/IP ports, drivers, and queues, so Intune deployments should install the SYSTEM sync scheduled task with `Install.ps1`.

By default queue names are prefixed with `PrintEdge - `, and stale cleanup only removes queues with that prefix or the PrintEdge management tag in the printer comment. Keep `removeUnmanagedPrinters` set to `false` unless you intentionally want the client to remove non-PrintEdge printers.

## Azure Driver Packages

Set `enableAzureDriverDownload` to `true` and configure `azureBlobBaseUrl`. Use `azureSasToken` when the container is private.

Printer or queue entries can specify:

- `driverBlobPath` - blob path under `azureBlobBaseUrl`.
- `driverUrl` - full URL, useful for per-driver SAS URLs.
- `infRelativePath` - path to the INF inside a downloaded zip.
- `infPath` - local INF path, relative to the script folder or absolute.

If the configured `driverName` is already installed, PrintEdge does not download or install the driver. Otherwise it tries the local `infPath`, then Azure Blob Storage when enabled. This applies to direct local queue mode. Print-server `connectionName` queues normally receive drivers from the print server.

Azure downloads are cached beside the app under `Drivers\<DriverName>`. For example, when the app is dropped into `C:\Program Files\PrintEdge`, a Kyocera driver downloads into `C:\Program Files\PrintEdge\Drivers\Kyocera TASKalfa 4004i KX`.

If Azure downloads are enabled while running as a standard user from `C:\Program Files\PrintEdge`, deploy a writable `Drivers` folder under the PrintEdge folder or pre-cache the driver packages there.

## Production Notes

- Replace dummy IPs, queues, and `Microsoft IPP Class Driver` with production printer IPs and vendor driver names.
- Use `Install.ps1` from Intune so direct local queue creation runs as SYSTEM.
- Test every driver package on the target Windows architecture before broad rollout.
- Host driver packages as `.zip` files containing the INF and related files, or direct `.inf` blobs.
- Prefer per-driver SAS URLs or managed distribution tooling over long-lived container SAS tokens.
- Keep `removeUnmanagedPrinters` disabled for pilot deployments.
