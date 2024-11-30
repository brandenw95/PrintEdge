# PrinterManager

**PrinterManager** is a Python-based application designed to manage network printers dynamically based on the current subnet. It automatically detects available printers, installs the necessary drivers, and removes printers that are no longer in use. The application also provides a system tray interface for ease of access.

---

## Features
- **Dynamic Printer Management**: Automatically updates installed printers based on the connected network subnet.
- **Driver Installation**: Installs printers using drivers stored in a predefined folder structure.
- **System Tray Integration**: A tray icon allows users to view available printers and exit the application conveniently.
- **GUI for Printers**: Displays a simple GUI listing the available printers for the current subnet.

---

## Prerequisites
1. **Python 3.7+**
2. Required Python libraries:
   - `pystray`
   - `Pillow`
   - `tkinter`
   - `pywin32`
3. Windows OS (tested with Windows 10/11).

---

## How to Use

### 1. Installation
- Clone the repository or download the source code.
- Ensure the required folder structure for printer drivers is in place.

### 2. Install Dependencies
Run the following command to install the required Python libraries:

pip install pystray pillow pywin32

### 3. Run the Application
Execute the script using:

python PrinterManager.py

### 4. System Tray Icon
- **Show Printers**: View the list of available printers for the current subnet.
- **Exit**: Exit the application.

---

## Customization
### Printer Mapping
Modify the `printer_mapping` dictionary in the `main()` function to include your network subnets and printer names:

printer_mapping = {
    "192.168.1.0": ["Takalfa 4002i", "HP LaserJet Pro"],
    "192.168.0.0": ["Canon Pixma MG2525", "Brother HL-L2350DW"],
}

### Driver Paths
Store the appropriate printer drivers in the corresponding folder under `C:\temp\printers`.

---

## Known Issues
1. **Driver Compatibility**: Ensure drivers are compatible with the system architecture (32-bit, 64-bit, ARM).
2. **Permissions**: Administrator privileges are required to install/remove printers.
3. **Icon Loading**: A default icon is used if `tray_icon.ico` is not found.

---

## Contributions
Contributions and suggestions are welcome. Please submit a pull request or create an issue for any bugs or feature requests.

