import os
import win32com.client
import tkinter as tk
from tkinter import filedialog

def create_startup_shortcut(target_exe_path, shortcut_name):
    # Path to the Startup folder
    startup_folder = os.path.join(os.environ['APPDATA'], 'Microsoft', 'Windows', 'Start Menu', 'Programs', 'Startup')
    
    # Full path for the shortcut
    shortcut_path = os.path.join(startup_folder, f"{shortcut_name}.lnk")
    
    # Create a COM object for the Shell
    shell = win32com.client.Dispatch('WScript.Shell')
    
    # Create the shortcut
    shortcut = shell.CreateShortCut(shortcut_path)
    shortcut.Targetpath = target_exe_path
    shortcut.WorkingDirectory = os.path.dirname(target_exe_path)
    shortcut.IconLocation = target_exe_path
    shortcut.save()
    
    print(f"Shortcut created at: {shortcut_path}")

def main():
    # Set up the Tkinter root
    root = tk.Tk()
    root.withdraw()  # Hide the main window
    
    # Open a file dialog to select the .exe file
    exe_path = filedialog.askopenfilename(
        title="Select the executable file",
        filetypes=[("Executable files", "*.exe")]
    )
    
    if not exe_path:
        print("No file selected.")
        return
    
    # Ask for the shortcut name
    shortcut_name = input("Enter the name for the shortcut: ")
    
    create_startup_shortcut(exe_path, shortcut_name)

if __name__ == "__main__":
    main()
