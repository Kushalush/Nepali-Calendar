import tkinter as tk
from datetime import datetime
from tkinter import font as tkinterFont
import os
import sys
import shutil
import winshell
from win32com.client import Dispatch

NEPALI_CALENDAR = {
    2082: {
        "Baisakh": (31, 2025, 4, 13),
        "Jestha": (31, 2025, 5, 14),
        "Ashad": (32, 2025, 6, 15),
        "Shrawan": (32, 2025, 7, 17),
        "Bhadra": (31, 2025, 8, 18),
        "Ashoj": (30, 2025, 9, 19),
        "Kartik": (30, 2025, 10, 19),
        "Mangshir": (29, 2025, 11, 18),
        "Poush": (30, 2025, 12, 18),
        "Magh": (29, 2026, 1, 17),
        "Falgun": (30, 2026, 2, 16),
        "Chaitra": (30, 2026, 3, 18)
    }
}

SETTINGS_FILE = "calendar_settings.txt"

def add_to_startup():
    script_path = os.path.abspath(sys.argv[0])
    startup_folder = winshell.startup()
    shortcut_path = os.path.join(startup_folder, "NepaliCalendar.lnk")

    if not os.path.exists(shortcut_path):
        shell = Dispatch('WScript.Shell')
        shortcut = shell.CreateShortCut(shortcut_path)
        shortcut.Targetpath = script_path
        shortcut.WorkingDirectory = os.path.dirname(script_path)
        shortcut.IconLocation = script_path
        shortcut.save()

class CompactCalendar:
    def __init__(self, root):
        self.root = root
        self.setup_window()
        self.setup_ui()
        self.update_dates()
        
    def setup_window(self):
        self.root.overrideredirect(True)
        self.root.attributes("-topmost", True)
        self.root.configure(bg="yellow", padx=5, pady=3)
        self.load_position()
        self.root.bind("<ButtonPress-1>", self.start_drag)
        self.root.bind("<B1-Motion>", self.on_drag)
        self.root.protocol("WM_DELETE_WINDOW", self.on_close)

    def setup_ui(self):
        title_font = tkinterFont.Font(family="Arial", size=10, weight="bold")
        date_font = tkinterFont.Font(family="Arial", size=9, weight="bold")

        self.english_date = tk.Label(self.root, bg="yellow", font=title_font)
        self.english_date.pack(pady=(2,0), anchor="w")
        
        self.nepali_date = tk.Label(self.root, bg="yellow", font=date_font)
        self.nepali_date.pack(anchor="w")

    def start_drag(self, event):
        self.drag_data = {"x": event.x, "y": event.y}

    def on_drag(self, event):
        new_x = self.root.winfo_x() + (event.x - self.drag_data["x"])
        new_y = self.root.winfo_y() + (event.y - self.drag_data["y"])
        self.root.geometry(f"+{new_x}+{new_y}")

    def convert_date(self, dt):
        for month, (days, y, m, d) in NEPALI_CALENDAR[2082].items():
            month_start = datetime(y, m, d)
            if dt >= month_start:
                delta = (dt - month_start).days
                if delta < days:
                    nep_month = ["बैशाख", "जेठ", "असार", "श्रावण", "भदौ", "असोज",
                                 "कार्तिक", "मंसिर", "पौष", "माघ", "फाल्गुन", "चैत्र"][list(NEPALI_CALENDAR[2082].keys()).index(month)]
                    return f"{nep_month} {delta+1} 2082"
        return "Date out of range"

    def update_dates(self):
        current_dt = datetime.now()
        self.english_date.config(text=current_dt.strftime("%B %d %Y").replace(" 0", " "))
        self.nepali_date.config(text=self.convert_date(current_dt))
        self.root.after(60000, self.update_dates)

    def save_position(self):
        pos = f"{self.root.winfo_x()},{self.root.winfo_y()}"
        with open(SETTINGS_FILE, "w") as f:
            f.write(pos)

    def load_position(self):
        if os.path.exists(SETTINGS_FILE):
            with open(SETTINGS_FILE, "r") as f:
                pos = f.read().strip().split(",")
                if len(pos) == 2:
                    try:
                        x, y = int(pos[0]), int(pos[1])
                        self.root.geometry(f"+{x}+{y}")
                        return
                    except:
                        pass
        self.root.geometry("400x300")

    def on_close(self):
        self.save_position()
        self.root.destroy()

if __name__ == "__main__":
    add_to_startup()  # ⚠️ only runs once
    root = tk.Tk()
    app = CompactCalendar(root)
    root.update_idletasks()
    width = app.english_date.winfo_reqwidth() + 40
    height = app.english_date.winfo_reqheight() * 2 + 25
    root.geometry(f"{width}x{height}+{root.winfo_x()}+{root.winfo_y()}")
    root.mainloop()
