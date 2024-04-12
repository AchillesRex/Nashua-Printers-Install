import os
import shutil
import subprocess
import tempfile
import sys
import ctypes

#Function for pop up message
def message_box(title, message):
    ctypes.windll.user32.MessageBoxW(0, message, title, 0)

def run_command(command, check=True):
    try:
        result = subprocess.run(command, shell=True, check=check)
        return result.returncode == 0
    except subprocess.CalledProcessError:
        return False

def printer_port_exists(port_name):
    command = f"powershell \"Get-PrinterPort -Name '{port_name}' 2>$null\""
    return run_command(command, check=False)

# Printer IP address
printer_ip = "<Secret>"

# Corrected network path to the driver's directory
network_driver_dir = r"<Secret>"

# Printer name
printer_name = "<Secret>"

# Driver name as specified in the .inf file
driver_name = "PCL6 Driver for Universal Print"

# Port name
port_name = f"IP_{printer_ip}"

# Create a temporary directory to copy the driver
with tempfile.TemporaryDirectory() as tmpdirname:
    try:
        # Copy the driver files to the temporary directory
        shutil.copytree(network_driver_dir, tmpdirname, dirs_exist_ok=True)

        # Path to the .inf file in the temporary directory
        local_driver_path = os.path.join(tmpdirname, "oemsetup.inf")

        # Check if the port already exists, and add it if it does not
        if not printer_port_exists(port_name):
            if not run_command(f"powershell Add-PrinterPort -Name \"{port_name}\" -PrinterHostAddress {printer_ip}"):
                print("Failed to add the printer port.")
                sys.exit(1)

        # Use pnputil to add the driver to the driver store
        if not run_command(f"pnputil.exe /add-driver \"{local_driver_path}\" /install"):
            print("Failed to add the driver to the driver store.")
            sys.exit(1)

        # Install the printer driver
        if not run_command(f"powershell Add-PrinterDriver -Name \\\"{driver_name}\\\""):
            print("Failed to install the printer driver.")
            sys.exit(1)


        # Install the printer using the TCP/IP port and the driver
        if not run_command(f"powershell Add-Printer -Name \"{printer_name}\" -DriverName \\\"{driver_name}\\\" -PortName \"{port_name}\""):
            print("Failed to install the printer.")
            sys.exit(1)


    except Exception as e:
        print(f"An error occurred: {e}")
        sys.exit(1)

# Remind user to set locked print code and paper size with a message box
message_title = "Printer Installation Successful"
message_body = f"Printer '{printer_name}' installed successfully.\nPlease remember to set the locked print code and default paper size to A4 in the printer settings."
message_box(message_title, message_body)
