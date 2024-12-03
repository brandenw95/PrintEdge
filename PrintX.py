import os
import platform
import socket
import threading
import tkinter as tk
from tkinter import ttk
import pystray
from pystray import MenuItem as item, Menu
from PIL import Image, ImageDraw
import win32print
import subprocess
import time

# --- Constants ---
BASE_DRIVER_PATH = "C:\\temp\\printers"

# --- Helper Functions ---
def get_current_subnet():
    """
    Detect the current subnet of the connected network.
    """
    hostname = socket.gethostname()
    ip_address = socket.gethostbyname(hostname)
    subnet = ".".join(ip_address.split(".")[:-1]) + ".0"
    return subnet

def list_printers_for_network(subnet, printer_mapping):
    """
    Return a list of printers based on the current subnet.
    """
    return printer_mapping.get(subnet, [])

def get_architecture_directory(printer_folder):
    """
    Determine the correct driver folder based on system architecture.
    """
    arch = platform.architecture()[0]
    if arch == "64bit":
        arch_folder = os.path.join(printer_folder, "64bit")
    elif arch == "32bit":
        arch_folder = os.path.join(printer_folder, "32bit")
    else:
        arch_folder = os.path.join(printer_folder, "arm64")
    
    # Return the folder if it exists; otherwise, fallback to the root folder
    if os.path.exists(arch_folder):
        return arch_folder
    else:
        return printer_folder

def install_printer(printer_name, driver_path):
    """
    Install the printer dynamically using the drivers from the specified path.
    """
    try:
        subprocess.run(
            ["rundll32", "printui.dll", "PrintUIEntry", "/if",
             f"/b{printer_name}", f"/f{driver_path}", "/r", f"//{printer_name}", "/m", "ModelName"],
            check=True
        )
    except Exception as e:
        print(f"Error installing printer {printer_name}: {e}")

def update_printers(subnet, printer_mapping):
    """
    Update the list of printers installed on the system based on the current subnet.
    """
    current_printers = list_printers_for_network(subnet, printer_mapping)
    installed_printers = [p[2] for p in win32print.EnumPrinters(2)]

    for printer in current_printers:
        if printer not in installed_printers:
            printer_folder = os.path.join(BASE_DRIVER_PATH, subnet, printer)
            driver_folder = get_architecture_directory(printer_folder)
            driver_path = os.path.join(driver_folder, "OEMSETUP.INF")
            if os.path.exists(driver_path):
                install_printer(printer, driver_path)
            else:
                print(f"Driver not found for printer {printer} in {driver_folder}")

    # Remove printers not in the current network
    for printer in installed_printers:
        if printer not in current_printers:
            try:
                subprocess.run(["rundll32", "printui.dll", "PrintUIEntry", "/dl",
                                f"/n{printer}"], check=True)
            except Exception as e:
                print(f"Error removing printer {printer}: {e}")

def create_image():
    """
    Load an .ico file from the current directory for the tray icon.
    """
    icon_path = os.path.join(os.path.dirname(__file__), "tray_icon.ico")
    try:
        image = Image.open(icon_path)
        return image
    except Exception as e:
        print(f"Error loading tray icon: {e}")
        # Fallback to a placeholder icon
        width = 64
        height = 64
        color1 = "black"
        color2 = "white"
        image = Image.new("RGB", (width, height), color1)
        dc = ImageDraw.Draw(image)
        dc.rectangle(
            (width // 4, height // 4, width * 3 // 4, height * 3 // 4),
            fill=color2,
        )
        return image

def show_printers_window(printers):
    """
    Display a simple GUI window showing available printers.
    """
    root = tk.Tk()
    root.title("Available Printers")
    ttk.Label(root, text="Available Printers").pack(pady=10)
    listbox = tk.Listbox(root, width=50, height=15)
    for printer in printers:
        listbox.insert(tk.END, printer)
    listbox.pack(pady=10)
    ttk.Button(root, text="Close", command=root.destroy).pack(pady=10)
    root.mainloop()

def check_and_create_folder_structure(printer_mapping):
    """
    Ensure that the required folder structure exists in BASE_DRIVER_PATH.
    """
    if not os.path.exists(BASE_DRIVER_PATH):
        os.makedirs(BASE_DRIVER_PATH)

    for subnet, printers in printer_mapping.items():
        subnet_path = os.path.join(BASE_DRIVER_PATH, subnet)
        if not os.path.exists(subnet_path):
            os.makedirs(subnet_path)

        for printer in printers:
            printer_path = os.path.join(subnet_path, printer)
            if not os.path.exists(printer_path):
                os.makedirs(printer_path)
                print(f"Created folder for printer: {printer_path}")

# --- Background Network Monitoring ---
def printer_manager_thread(printer_mapping):
    """
    Continuously check the network and update printers.
    """
    previous_subnet = None
    while True:
        current_subnet = get_current_subnet()
        if current_subnet != previous_subnet:
            print(f"Detected new subnet: {current_subnet}")
            update_printers(current_subnet, printer_mapping)
            previous_subnet = current_subnet
        time.sleep(10)  # Check every 10 seconds

# --- Tray Icon Setup ---
def setup_tray_icon(printer_mapping):
    """
    Setup the system tray icon and functionality.
    """
    def show_printers_action():
        subnet = get_current_subnet()
        printers = list_printers_for_network(subnet, printer_mapping)
        show_printers_window(printers)

    icon = pystray.Icon(
        "PrinterManager",
        create_image(),
        menu=Menu(
            item("Show Printers", lambda: show_printers_action()),
            item("Exit", lambda: exit(0)),
        ),
    )
    return icon

# --- Main Program ---
def main():
    # Define printer mappings for each subnet
    printer_mapping = {
        "192.168.1.0": ["Takalfa 4002i", "HP LaserJet Pro"],
        "192.168.0.0": ["Canon Pixma MG2525", "Brother HL-L2350DW"],
        "10.0.1.0": ["Kyocera 4004i", "Brother HL-L2350DW"],
    }

    # Check and create folder structure
    check_and_create_folder_structure(printer_mapping)

    # Start the printer management thread
    threading.Thread(target=printer_manager_thread, args=(printer_mapping,), daemon=True).start()

    # Setup the tray icon
    icon = setup_tray_icon(printer_mapping)
    icon.run()

if __name__ == "__main__":
    main()
