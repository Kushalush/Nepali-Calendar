import os
import shutil
import sys
import winshell
from win32com.client import Dispatch
import subprocess

APP_NAME = "NepaliCalendar"
INSTALL_DIR = os.path.join(os.environ["APPDATA"], APP_NAME)
SCRIPT_NAME = os.path.join(os.path.dirname(sys.executable), "calendar.py")
SHORTCUT_NAME = f"{APP_NAME}.lnk"

def copy_to_install_dir():
    os.makedirs(INSTALL_DIR, exist_ok=True)
    shutil.copy(SCRIPT_NAME, os.path.join(INSTALL_DIR, SCRIPT_NAME))

def create_startup_shortcut():
    startup_path = winshell.startup()
    shortcut_path = os.path.join(startup_path, SHORTCUT_NAME)
    target_path = os.path.join(INSTALL_DIR, SCRIPT_NAME)

    if not os.path.exists(shortcut_path):
        shell = Dispatch('WScript.Shell')
        shortcut = shell.CreateShortCut(shortcut_path)
        shortcut.Targetpath = sys.executable
        shortcut.Arguments = f'"{target_path}"'
        shortcut.WorkingDirectory = INSTALL_DIR
        shortcut.IconLocation = sys.executable
        shortcut.save()

def run_calendar():
    script_path = os.path.join(INSTALL_DIR, SCRIPT_NAME)
    subprocess.Popen([sys.executable, script_path], cwd=INSTALL_DIR)

def main():
    print("Installing Nepali Calendar...")
    copy_to_install_dir()
    create_startup_shortcut()
    print("Running calendar...")
    run_calendar()
    print("✅ Installed successfully.")

if __name__ == "__main__":
    main()
