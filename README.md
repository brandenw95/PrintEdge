# PrintEdge

PrintEdge is a Windows PowerShell tray application for internal printer mapping. It detects the active IPv4 subnet, selects a matching profile from `PrintEdge.config.json`, and adds or removes the print queues assigned to that subnet.

This project is intended as a lightweight internal Printix-style client. The included configuration uses dummy printers and the built-in `Microsoft IPP Class Driver` so the app can be validated without vendor driver packages.

## Features

- Subnet-to-printer and subnet-to-queue mapping through `PrintEdge.config.json`.
- CIDR matching, for example `192.168.1.0/24`.
- Generated system tray icon with active subnet status, manual sync, and quit.
- Double-click tray popup showing only the active subnet's printers and queues.
- Safe cleanup that removes only PrintEdge-managed queues by default.
- Non-admin runtime when queues use `connectionName` print-server shares.
- Optional Azure Blob Storage driver download for `.zip` or `.inf` driver packages.
- CLI validation and one-shot sync modes for deployment automation.
- Log output to `%ProgramData%\PrintEdge\Logs\PrintEdge.log` by default.

## Files

- `PrintX.ps1` - main PrintEdge tray and sync program.
- `PrintEdge.config.json` - subnet, printer, queue, driver, and Azure settings.

## Requirements

- Windows 10/11 or Windows Server with the Print Management PowerShell cmdlets.
- Windows PowerShell 5.1.
- No administrator rights are required for normal runtime when queues use `connectionName`.
- Direct local TCP/IP queue creation and system driver installation are Windows admin operations. PrintEdge skips those gracefully when it is not elevated.

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

Create `C:\Program Files\PrintEdge`, then place these two files in that folder:

- `PrintX.ps1`
- `PrintEdge.config.json`

Validate the dropped files:

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "C:\Program Files\PrintEdge\PrintX.ps1" -ValidateConfig
```

Run the tray app from that location:

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File "C:\Program Files\PrintEdge\PrintX.ps1"
```

If your deployment tool creates a scheduled task, startup item, or shortcut, point it at the same command. Use print-server `connectionName` queues for non-admin runtime.

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
      "connectionName": "\\\\printserver\\HQ-Color-MFP",
      "comment": "Default color queue"
    }
  ]
}
```

Use `connectionName` for production non-admin mapping. Without `connectionName`, PrintEdge treats the entry as a direct local TCP/IP queue, which Windows requires elevation to create.

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
- Use `connectionName` print-server shares for a fully non-admin client experience.
- Test every driver package on the target Windows architecture before broad rollout.
- Host driver packages as `.zip` files containing the INF and related files, or direct `.inf` blobs.
- Prefer per-driver SAS URLs or managed distribution tooling over long-lived container SAS tokens.
- Keep `removeUnmanagedPrinters` disabled for pilot deployments.
