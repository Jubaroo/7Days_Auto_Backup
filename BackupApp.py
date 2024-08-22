import logging
import os
import tkinter as tk
from tkinter import ttk
import zipfile
from datetime import datetime, timedelta
from threading import Timer, Thread
from tkinter import filedialog, messagebox
import pystray
from PIL import Image, ImageDraw
import win32api


def validate_numeric(value):
    if value.isdigit() or value == "":
        return True
    return False

def center_window(window):
    window.update_idletasks()
    window_width = window.winfo_width()
    window_height = window.winfo_height()
    screen_width = window.winfo_screenwidth()
    screen_height = window.winfo_screenheight()
    x = (screen_width // 2) - (window_width // 2)
    y = (screen_height // 2) - (window_height // 2)
    window.geometry(f"+{x}+{y}")

class BackupApp:
    def __init__(self, master):
        self.root = master
        self.root.title("7 Days to Die Auto Backup")
        self.root.withdraw()  # Hide the main window initially
        self.server_folder = tk.StringVar()
        self.backup_folder = tk.StringVar()
        self.backup_mode = tk.StringVar(value="interval")
        self.backup_interval = tk.IntVar(value=15)
        self.backup_hour = tk.IntVar(value=2)
        self.backup_minute = tk.IntVar(value=0)
        self.am_pm = tk.StringVar(value="AM")
        self.max_backups = tk.IntVar(value=10)
        self.create_widgets()
        center_window(master)

        self.backup_timer = None
        self.icon_thread = None  # Thread for the system tray icon

        self.setup_tray_icon()  # Setup system tray icon

    def setup_tray_icon(self):
        def run_icon():
            image = Image.new('RGB', (64, 64), color="black")
            draw = ImageDraw.Draw(image)
            draw.rectangle([16, 16, 48, 48], outline="white", fill="gray")

            icon = pystray.Icon("backup_app", image, "7 Days to Die Auto Backup", menu=pystray.Menu(
                pystray.MenuItem("Show", self.show_window),
                pystray.MenuItem("Exit", self.exit_app)
            ))
            self.icon = icon
            icon.run()

        # Start the system tray icon in a separate thread
        self.icon_thread = Thread(target=run_icon, daemon=True)
        self.icon_thread.start()

    def show_window(self, icon=None, item=None):
        self.root.deiconify()  # Show the main window
        self.root.update_idletasks()  # Force the window to update and redraw
        self.root.update()  # Ensure the layout is fully recalculated
        self.root.lift()  # Bring the window to the front
        self.root.focus_force()  # Focus on the window

    def exit_app(self, icon=None, item=None):
        if self.backup_timer:
            self.backup_timer.cancel()
        if self.icon:
            self.icon.stop()
        self.root.destroy()  # Destroy the main window and exit the app

    def create_widgets(self):
        validate_command = self.root.register(validate_numeric)

        # Game Data Folder with Browse and Auto-Find buttons
        tk.Label(self.root, text="Game Data Folder:").grid(row=0, column=0, padx=10, pady=10, sticky="e")
        tk.Entry(self.root, textvariable=self.server_folder, width=50).grid(row=0, column=1, padx=(0, 10), pady=10, sticky="we")
        tk.Button(self.root, text="Browse", command=self.browse_server_folder).grid(row=0, column=2, padx=(10, 5), pady=10)
        tk.Button(self.root, text="Auto-Find", command=self.auto_find_game_data).grid(row=0, column=3, padx=(0, 10), pady=10)

        # Backup Folder
        tk.Label(self.root, text="Backup Folder:").grid(row=1, column=0, padx=10, pady=10, sticky="e")
        tk.Entry(self.root, textvariable=self.backup_folder, width=50).grid(row=1, column=1, padx=(0, 10), pady=10, sticky="we")
        tk.Button(self.root, text="Browse", command=self.browse_backup_folder).grid(row=1, column=2, padx=(10, 10), pady=10)

        # Backup Mode
        tk.Label(self.root, text="Backup Mode:").grid(row=2, column=0, padx=10, pady=10, sticky="e")
        mode_frame = tk.Frame(self.root)
        mode_frame.grid(row=2, column=1, columnspan=2, pady=10, sticky="w")
        tk.Radiobutton(mode_frame, text="Interval", variable=self.backup_mode, value="interval",
                       command=self.toggle_backup_mode).pack(side=tk.LEFT)
        tk.Radiobutton(mode_frame, text="Time of Day", variable=self.backup_mode, value="time_of_day",
                       command=self.toggle_backup_mode).pack(side=tk.LEFT)

        # Backup Interval and Max Backups in the same row, closer together
        interval_backups_frame = tk.Frame(self.root)
        interval_backups_frame.grid(row=3, column=0, columnspan=4, pady=10, sticky="w")

        tk.Label(interval_backups_frame, text="Backup Interval (minutes):").pack(side=tk.LEFT, padx=(10, 5))
        tk.Entry(interval_backups_frame, textvariable=self.backup_interval, width=10, validate="key", validatecommand=(validate_command, '%P')).pack(side=tk.LEFT, padx=(0, 20))

        tk.Label(interval_backups_frame, text="Max Backups:").pack(side=tk.LEFT, padx=(0, 5))
        tk.Entry(interval_backups_frame, textvariable=self.max_backups, width=10, validate="key", validatecommand=(validate_command, '%P')).pack(side=tk.LEFT)

        # Time of Day (remains below the interval/max backups row)
        self.time_frame = tk.Frame(self.root)
        tk.Label(self.time_frame, text="Backup Time:").grid(row=0, column=0, sticky="e")
        tk.Entry(self.time_frame, textvariable=self.backup_hour, width=3, validate="key", validatecommand=(validate_command, '%P')).grid(row=0, column=1)
        tk.Label(self.time_frame, text=":").grid(row=0, column=2)
        tk.Entry(self.time_frame, textvariable=self.backup_minute, width=3, validate="key", validatecommand=(validate_command, '%P')).grid(row=0, column=3)
        tk.Checkbutton(self.time_frame, text="AM", variable=self.am_pm, onvalue="AM", offvalue="PM").grid(row=0, column=4)
        tk.Checkbutton(self.time_frame, text="PM", variable=self.am_pm, onvalue="PM", offvalue="AM").grid(row=0, column=5)
        self.time_frame.grid(row=4, column=0, columnspan=4, pady=10, sticky="w")

        # Start Backup Button
        tk.Button(self.root, text="Start Backup", command=self.start_backup).grid(row=5, column=0, columnspan=4, pady=10)

        # Progress Bar
        self.progress = ttk.Progressbar(self.root, orient="horizontal", length=300, mode="determinate")
        self.progress.grid(row=6, column=0, columnspan=4, pady=10)

        # Status Label
        self.status_label = tk.Label(self.root, text="", fg="blue")
        self.status_label.grid(row=7, column=0, columnspan=4, pady=10)

        self.toggle_backup_mode()

        # Make columns 1 and 3 expand to fill available space
        self.root.grid_columnconfigure(1, weight=1)
        self.root.grid_columnconfigure(3, weight=1)

    def auto_find_game_data(self):
        possible_dirs = []
        drives = win32api.GetLogicalDriveStrings().split('\000')[:-1]
        steam_paths = ["Program Files (x86)", "Program Files", ""]

        for drive in drives:
            for path in steam_paths:
                steam_library_path = os.path.join(drive, path, "SteamLibrary", "steamapps", "common",
                                                  "7 Days to Die Dedicated Server", "saves")
                if os.path.exists(steam_library_path):
                    possible_dirs.append(steam_library_path)

        if not possible_dirs:
            messagebox.showerror("Error", "Could not find the 7 Days to Die Dedicated Server saves directory.")
            return

        latest_dir = max(possible_dirs, key=lambda d: os.path.getmtime(d))

        # Check if the saves folder contains any subdirectories
        save_folders = [os.path.join(latest_dir, d) for d in os.listdir(latest_dir) if
                        os.path.isdir(os.path.join(latest_dir, d))]
        if not save_folders:
            messagebox.showerror("Error", "The saves directory does not contain any folders.")
            return

        # Find the most recently modified folder in the saves directory
        most_recent_save = max(save_folders, key=os.path.getmtime)
        self.server_folder.set(most_recent_save)

    def browse_server_folder(self):
        folder = filedialog.askdirectory()
        if folder:
            self.server_folder.set(folder)

    def browse_backup_folder(self):
        folder = filedialog.askdirectory()
        if folder:
            self.backup_folder.set(folder)

    def start_backup(self):
        if not os.path.isdir(self.server_folder.get()):
            messagebox.showerror("Error", "Invalid server folder.")
            self.status_label.config(text="Error: Invalid server folder.", fg="red")
            logging.error("Invalid server folder selected.")
            return

        if not os.path.isdir(self.backup_folder.get()):
            messagebox.showerror("Error", "Invalid backup folder.")
            self.status_label.config(text="Error: Invalid backup folder.", fg="red")
            logging.error("Invalid backup folder selected.")
            return

        # Perform an immediate backup in a separate thread
        self.status_label.config(text="Starting immediate backup...", fg="green")
        backup_thread = Thread(target=self.backup)
        backup_thread.start()

    def backup(self):
        try:
            self.progress["value"] = 0
            self.status_label.config(text="Backing up...", fg="green")
            server_folder = self.server_folder.get()
            backup_folder = self.backup_folder.get()
            timestamp = datetime.now().strftime("%Y%m%d_%H%M")
            backup_filename = f"7d2d_backup_{timestamp}.zip"
            backup_path = os.path.join(backup_folder, backup_filename)

            # Calculate total number of files for progress tracking
            total_files = sum([len(files) for r, d, files in os.walk(server_folder)])
            processed_files = 0

            with zipfile.ZipFile(backup_path, 'w') as backup_zip:
                for foldername, subfolders, filenames in os.walk(server_folder):
                    for filename in filenames:
                        file_path = os.path.join(foldername, filename)
                        backup_zip.write(file_path, os.path.relpath(file_path, server_folder))

                        # Update progress
                        processed_files += 1
                        progress = (processed_files / total_files) * 100
                        self.progress["value"] = progress
                        self.status_label.config(text=f"Backing up... {processed_files}/{total_files} files", fg="green")
                        self.root.update_idletasks()

            self.status_label.config(text=f"Backup completed: {backup_filename}", fg="green")
            logging.info(f"Backup created: {backup_filename}")
            self.rotate_backups(backup_folder)

            # Schedule the next backup
            next_backup_time = self.schedule_backup()
            self.status_label.config(text=f"Next backup scheduled at {next_backup_time}", fg="blue")
        except Exception as e:
            self.status_label.config(text="Backup failed!", fg="red")
            logging.error(f"Backup failed: {str(e)}")

    def schedule_backup(self):
        if self.backup_timer:
            self.backup_timer.cancel()

        if self.backup_mode.get() == "interval":
            interval = self.backup_interval.get() * 60
            next_backup_time = datetime.now() + timedelta(seconds=interval)
            self.backup_timer = Timer(interval, self.backup)
        else:
            hour = self.backup_hour.get()
            minute = self.backup_minute.get()
            am_pm = self.am_pm.get()
            if am_pm == "PM" and hour < 12:
                hour += 12
            elif am_pm == "AM" and hour == 12:
                hour = 0
            now = datetime.now()
            backup_time = now.replace(hour=hour, minute=minute, second=0, microsecond=0)
            if now > backup_time:
                backup_time += timedelta(days=1)
            delay = (backup_time - now).total_seconds()
            next_backup_time = backup_time
            self.backup_timer = Timer(delay, self.backup)

        self.backup_timer.start()
        return next_backup_time.strftime('%Y-%m-%d %H:%M:%S')

    def rotate_backups(self, folder):
        backups = sorted([f for f in os.listdir(folder) if f.startswith("7d2d_backup_")], reverse=True)
        if len(backups) > self.max_backups.get():
            for old_backup in backups[self.max_backups.get():]:
                os.remove(os.path.join(folder, old_backup))

    def stop_backup(self):
        if self.backup_timer:
            self.backup_timer.cancel()

    def toggle_backup_mode(self):
        if self.backup_mode.get() == "interval":
            # The interval/max backups row is already visible, so no action needed
            self.time_frame.grid_forget()  # Hide time-of-day selection
        else:
            # Show the time-of-day selection
            self.time_frame.grid(row=4, column=0, columnspan=4, pady=10, sticky="w")
            # No need to hide the interval/max backups, they can remain visible

if __name__ == "__main__":
    root = tk.Tk()
    app = BackupApp(root)
    root.protocol("WM_DELETE_WINDOW", app.exit_app)
    root.mainloop()
