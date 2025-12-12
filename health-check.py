import os
import sys
import subprocess
import ctypes
import webbrowser
import psutil
import tkinter as tk
from tkinter import ttk
from ttkthemes import ThemedTk
from tkinter import messagebox
import emoji
import get_passcode_v2

# Status symbols
goodStatus = emoji.emojize(':thumbs_up:')
badStatus = emoji.emojize(':thumbs_down:')
batteryStatus = emoji.emojize(':battery:')
thermometer = emoji.emojize(':face_with_thermometer:')
username = ""

# ----------- Get current username -------------------#
# ----------- Functional, but not being used. Will need to create Tkinter GUI -------------------#
#def GetUserName():
#    username = entry_user.get().lower()
#    login_window.destroy()
#    print(username)
#

# ----------- Check CPU and Memory usage -------------#


def CheckCPUandRAM():
    # Utility variables
    cpuUsage = psutil.cpu_percent(interval=1)
    memoryUsage = psutil.virtual_memory().percent
    memoryTotal = round(psutil.virtual_memory().total / 1000000000, 0)
    
    # Prevents multiple execution on first click
    cpuAndMemoryBtn.config(state=tk.DISABLED)

    # Checks for good or bad resource usage
    if cpuUsage > 90 or memoryUsage > 80:
        cpuAndMemoryOverview.config(text=f"{badStatus} High resource usage! {badStatus}")
        cpuLabel.config(text=f"CPU Usage: {cpuUsage}%")
        memoryLabel.config(text=f"Memory Usage: {memoryUsage}%")
        print(f"{badStatus} High resource usage! {badStatus} | CPU: {cpuUsage}% | Memory: {memoryUsage}%")
    else: 
        cpuAndMemoryOverview.config(text=f"{goodStatus} Good resource usage! {goodStatus}")
        cpuLabel.config(text=f"CPU Usage: {cpuUsage}%")
        memoryLabel.config(text=f"Memory Usage: {memoryUsage}%")
        print(f"{goodStatus} Good resource usage! {goodStatus} | CPU: {cpuUsage}% | Memory: {memoryUsage}%")

    totalMemoryLabel.config(text=f"Total Device Memory: {memoryTotal} GB\n")

    # Enables button to run execution if clicked again
    cpuAndMemoryBtn.config(state=tk.NORMAL)


# ----------- Generate battery condition report -------------#


def CheckBatteryCondition():
    # Utility variable
    outputPath = os.path.expanduser("~/Downloads/battery-report.html")

    # Pull battery report
    try:
        subprocess.run(["powercfg", "/batteryreport", f"/output", outputPath], check=True)
        print(f"\n{batteryStatus} Battery report generated: {outputPath}\n")
        webbrowser.open(outputPath)

    # Error catch
    except subprocess.CalledProcessError as e:
        print(f"\n{badStatus} Error generating battery report:\n", e)


# ----------- Run HP Driver Program -------------#


def CheckHPDrivers():
    # If user is elevated, run HP driver application as admin
    if CheckElevate():
        def resource_path(relative_path):
            try:
                base_path = sys._MEIPASS
            except Exception:
                base_path = os.path.abspath(".")

            return os.path.join(base_path, relative_path)
        
        def run_as_admin(exe_path):
            try:
                ctypes.windll.shell32.ShellExecuteW(None, "runas", exe_path, None, None, 1)
            except Exception as e:
                print("Failed to launch with admin privileges:", e)
        
        external_app = resource_path("hp-hpia-5.3.2.exe")
        run_as_admin(external_app)

    # If user is NOT elevated, run elevate command in terminal
    else:
        proc = subprocess.Popen(
            ['cmd', '/k', 'elevate'],
            creationflags=0x00000010
        )


# ----------- Run Lenovo Driver Program -------------#


def CheckLenovoDrivers():
    # If user is elevated, run Lenovo driver application as admin
    if CheckElevate():
        def resource_path(relative_path):
            try:
                base_path = sys._MEIPASS
            except Exception:
                base_path = os.path.abspath(".")

            return os.path.join(base_path, relative_path)
        
        def run_as_admin(exe_path):
            try:
                ctypes.windll.shell32.ShellExecuteW(None, "runas", exe_path, None, None, 1)
            except Exception as e:
                print("Failed to launch with admin privileges:", e)
        
        external_app = resource_path("system_update_5.08.03.59.exe")
        run_as_admin(external_app)

    # If user is NOT elevated, run elevate command in terminal
    else:
        proc = subprocess.Popen(
            ['cmd', '/k', 'elevate'],
            creationflags=0x00000010
        )

        
# ----------- Run Windows 11 Search UI repair -------------#


def SearchUIRepair():
    # If user is elevated, run run Search UI repair commands in elevated Powershell
    if CheckElevate():
        def run_as_admin():
            try:
                search_repair_cmd = (
                    'Start-Process powershell -Verb runAs '
                    '-ArgumentList \'-NoProfile -Command "'
                    'Get-Service -Name WSearch | Restart-Service; '
                    'taskkill /f /im SearchHost.exe'
                    '"\''
                )

                subprocess.run(["powershell", "-Command", search_repair_cmd])
                print("Search UI repaired!")
            except Exception as e:
                print("Failed to launch with admin privileges:", e)

        run_as_admin()

    # If user is NOT elevated, run elevate command in terminal
    else:
        proc = subprocess.Popen(
            ['cmd', '/k', 'elevate'],
            creationflags=0x00000010
        )


# ----------- Delete Copied ZDesigner printers -------------#
# -------------------- NON_FUNCTIONAL ----------------------#


#def DeleteClonedPrinters():
#    # If user is elevated, run run Search UI repair commands in elevated Powershell
#    if CheckElevate():
#        def run_as_admin():
#            try:
#                delete_printers_cmd = (
#                    'Start-Process powershell -Verb RunAs -ArgumentList '
#                    '"-NoProfile -Command '
#                    "Get-WmiObject -Query \\\"SELECT * FROM Win32_Printer WHERE Name LIKE 'ZDesigner%'\\\" | ForEach-Object { $_.Delete() }"
#                    '"'
#                )
#               subprocess.run(["powershell", "-Command", delete_printers_cmd])
#                print("ZDesigner printers deleted!")
#            except Exception as e:
#                print("Failed to launch with admin privileges:", e)
#
#        run_as_admin()
#
#    # If user is NOT elevated, run elevate command in terminal
#    else:
#        proc = subprocess.Popen(
#            ['cmd', '/k', 'elevate'],
#            creationflags=0x00000010
#        )

        
# ----------- Get MC33/TC56/TC57 Admin password -------------------#


def GetZebraAdmin():
    serialNumber = serial_entry.get()
    print(serialNumber)
    selectedRegion = regionDropdown.get()
    
    if getattr(sys, 'frozen', False) and hasattr(sys, '_MEIPASS'):
        # Running in a PyInstaller bundle
        base_path = sys._MEIPASS
    else:
        # Running in normal Python
        base_path = os.path.dirname(__file__)

    script_path = os.path.join(base_path, "get_passcode_v2.py")

    try:
        proc = subprocess.Popen(
            ["python", script_path, "-s", serialNumber, "-r", selectedRegion],
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            text=True
        ) 
        
        for line in proc.stdout:
            print(line.strip())

        if any(adminPW in line for adminPW in ["Invalid login token."]):
            print("Midway issue.")
            adminPasswordLabel.config(text="Failed to get Midway token", fg="red")
            proc = subprocess.Popen(
                ['cmd', '/k', 'mwinit', '--aea'],
                creationflags=0x00000010
                ) 
            
        if any(adminPW in line for adminPW in ["404 Client Error:"]):
            print("Incorrect region.")
            adminPasswordLabel.config(text="404 Client Error: Incorrect region.", fg="red")
        
        if any(adminPW in line for adminPW in ["Passcode:"]):
            print("Admin Password Found!")
            adminPasswordLabel.config(text=line, fg="green")
            

    # Error catch
    except subprocess.CalledProcessError as e:
        adminPasswordLabel.config(text="Error generating Admin Password.", fg="red")
        print("Error generating Admin Password.", e)
 

# ----------- Check if USER is elevated -------------#


def CheckElevate():
    elevated = False
    username = os.environ.get("USERNAME") or os.environ.get("USER")
    # Check local admin on device
    proc = subprocess.Popen(
        ['net', 'localgroup', 'administrators'],
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        text=True
    )

    for line in proc.stdout:
        print(line.strip())
        
        if any(adminUser in line for adminUser in [f"ANT\\{username}-Admin", f"ANT\\{username}-admin"]) or "-admin" in username.lower():
            elevated = True
            print("User is Elevated!")
            elevateLabel.config(text="Is user elevated? Yes", fg="green")
            break
    
    if elevated == False:
        print("User not elevated. Please elevate before proceeding.")
        elevateLabel.config(text="Is user elevated? No.\nOnce elevated, please try selecting one of the choice below, again.", fg="red")
    
    return elevated
        
        
# Handle file paths correctly when running as .exe
if getattr(sys, 'frozen', False):
    basedir = sys._MEIPASS
else:
    basedir = os.path.dirname(os.path.abspath(__file__))

# Load Emoji JSON correctly
emoji_path = os.path.join(basedir, "emoji", "unicode_codes", "emoji.json")
icon_path = os.path.join(basedir, "favicon.ico")


# ----------- User Interface -------------#


# GUI 
# layout and styles
app = ThemedTk(theme="ubuntu")
app.iconbitmap(icon_path)
app.title("System Health Report")
app.geometry("600x700")
style = ttk.Style()
style.configure("TButton", font=("Helvetica", 12), margin=5)

# GUI 
# title
headerLabel = tk.Label(app, text=f"System Health Report {thermometer}", font=("Comic Sans MS", 28))
headerLabel.pack(side="top")

# GUI 
# CPU and Memory section
cpuAndMemoryOverview = tk.Label(app)
cpuAndMemoryOverview.pack()
cpuLabel = tk.Label(app, text="CPU Usage: ", font=("Helvetica", 14))
cpuLabel.pack(anchor="center")
memoryLabel = tk.Label(app, text="Memory Usage: ", font=("Helvetica", 14))
memoryLabel.pack()
totalMemoryLabel = tk.Label(app, text="Total Device Memory: \n", font=("Helvetica", 14))
totalMemoryLabel.pack()
cpuAndMemoryBtn = ttk.Button(app, text="Run system health check", command=CheckCPUandRAM, style="TButton")
cpuAndMemoryBtn.pack()

# GUI
# Battery report button
batteryBtn = ttk.Button(app, text="Generate battery condition report", command=CheckBatteryCondition, style="TButton")
batteryBtn.pack(pady=5)

# GUI
# HP Drivers button
elevateLabel = tk.Label(app, text="", font=("Helvetica", 14))
elevateLabel.pack(anchor="center")
hpDriverBtn = ttk.Button(app, text="Check for HP Driver updates", command=CheckHPDrivers, style="TButton")
hpDriverBtn.pack(pady=5)

# GUI
# Lenovo Drivers button
lenovoDriverBtn = ttk.Button(app, text="Check for Lenovo Driver updates", command=CheckLenovoDrivers, style="TButton")
lenovoDriverBtn.pack(pady=5)

# GUI
# Search UI Repair button
searchUIBtn = ttk.Button(app, text="Repair Windows 11 Search UI", command=SearchUIRepair, style="TButton")
searchUIBtn.pack(pady=5)

# GUI - Goes with "Delete Copied ZDesigner printers" function (NON_FUNCTIONAL)
# Delete cloned printers button
# deletePrintersBtn = ttk.Button(app, text="Delete ZDesigner Cloned Printers", command=DeleteClonedPrinters, style="TButton")
# deletePrintersBtn.pack(pady=5)

# GUI
# Get MC33/TC56/TC57 Admin password
tk.Label(app, text="\n\nEnter MC33/TC56/TC57 Serial Number:", font=("Helvetica", 12)).pack()
serial_entry = tk.Entry(app)
serial_entry.pack()
tk.Label(app, text="Select your region:", font=("Helvetica", 10)).pack()
regionDropdown = ttk.Combobox(app, values=["us-east-1", "eu-west-1", "ap-south-1", "ap-northeast-1"])
regionDropdown.current(0)
regionDropdown.pack()
regionDropdown.bind("<<ComboboxSelected>>", GetZebraAdmin)
passwordBtn = ttk.Button(app, text="Check for Admin Password", command=GetZebraAdmin, style="TButton")
passwordBtn.pack()
adminPasswordLabel = tk.Label(app, text="", font=("Helvetica", 14))
adminPasswordLabel.pack(anchor="center")


# GUI
# Footer
footerLabel = tk.Label(app, text="Version 4.0 | Collin Dapper", font=("Comic Sans MS", 10))
footerLabel.pack(side="bottom")

app.mainloop()
