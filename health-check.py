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



# ----------- Run Lenovo Driver Program -------------#


def CheckLenovoDrivers():
    # If user is elevated, run Lenovo driver application as admin
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

        
# ----------- Run Windows 11 Search UI repair -------------#


def SearchUIRepair():
    # If user is elevated, run run Search UI repair commands in elevated Powershell
    
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
cpuAndMemoryBtn.pack(pady=5)

# GUI
# Battery report button
batteryBtn = ttk.Button(app, text="Generate battery condition report", command=CheckBatteryCondition, style="TButton")
batteryBtn.pack(pady=5)

# GUI
# HP Drivers button
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
# Footer
footerLabel = tk.Label(app, text="Version 4.0 | Collin Dapper", font=("Comic Sans MS", 10))
footerLabel.pack(side="bottom")

app.mainloop()
