import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import requests
import os
import threading
from pathlib import Path
import queue


class MP3DownloaderApp:
    def __init__(self, root):
        self.root = root
        self.root.title("MP3 Downloader")
        self.root.geometry("700x550")
        self.root.resizable(False, False)
        
        # Center the window on the screen
        self.center_window()

        # Set background color for the entire window
        self.root.configure(bg="white")

        # UI Elements
        self.file_label = tk.Label(root, text="Select an Excel File", font=("Helvetica", 12), bg="white", fg="black")
        self.file_label.pack(pady=10)

        self.select_file_button = tk.Button(root, text="Browse File", command=self.select_file, bg="lightgray", fg="black")
        self.select_file_button.pack(pady=5)

        self.download_dir_button = tk.Button(root, text="Choose Download Directory", command=self.choose_download_directory, bg="lightgray", fg="black")
        self.download_dir_button.pack(pady=5)

        self.start_button = tk.Button(root, text="Start Download", command=self.start_download_thread, state=tk.DISABLED, bg="lightgray", fg="black")
        self.start_button.pack(pady=5)

        self.pause_button = tk.Button(root, text="Pause", command=self.toggle_pause, state=tk.DISABLED, bg="lightgray", fg="black")
        self.pause_button.pack(pady=5)

        self.stop_button = tk.Button(root, text="Stop", command=self.stop_download, state=tk.DISABLED, bg="lightgray", fg="black")
        self.stop_button.pack(pady=5)

        self.log_label = tk.Label(root, text="Download Logs:", font=("Helvetica", 12), bg="white", fg="black")
        self.log_label.pack(pady=5)

        # Log Frame with Scrollbar
        self.log_frame = tk.Frame(root, bg="white")
        self.log_frame.pack(pady=10)

        self.log_scrollbar = tk.Scrollbar(self.log_frame)
        self.log_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        self.log_text = tk.Text(
            self.log_frame,
            height=20,
            width=80,
            bg="white",
            fg="black",
            font=("Helvetica", 10),
            state=tk.DISABLED,
            yscrollcommand=self.log_scrollbar.set
        )
        self.log_text.pack(side=tk.LEFT, fill=tk.BOTH)

        self.log_scrollbar.config(command=self.log_text.yview)

        # Variables
        self.excel_file_path = None
        self.download_directory = "downloads"  # Default download directory
        self.download_log_path = "download_log.csv"
        self.is_paused = False
        self.is_downloading = False
        self.stop_thread = False
        self.log_queue = queue.Queue()

        self.update_logs()  # Start log update loop

    def center_window(self):
        """Center the app window on the screen."""
        self.root.update_idletasks()
        width = self.root.winfo_width()
        height = self.root.winfo_height()
        x = (self.root.winfo_screenwidth() // 2) - (width // 2)
        y = (self.root.winfo_screenheight() // 2) - (height // 2)
        self.root.geometry(f"{width}x{height}+{x}+{y}")

    def select_file(self):
        file_path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx *.xls")])
        if file_path:
            self.excel_file_path = file_path
            self.file_label.config(text=f"Selected File: {os.path.basename(file_path)}")
            self.start_button.config(state=tk.NORMAL)

    def choose_download_directory(self):
        """Allow the user to choose a directory for downloads."""
        directory = filedialog.askdirectory()
        if directory:
            self.download_directory = directory
            self.log_message(f"Download directory set to: {self.download_directory}")

    def log_message(self, message):
        """Add log messages to the queue."""
        self.log_queue.put(message)

    def update_logs(self):
        """Process log messages from the queue."""
        while not self.log_queue.empty():
            message = self.log_queue.get()
            self.log_text.config(state=tk.NORMAL)
            self.log_text.insert(tk.END, message + "\n")
            self.log_text.see(tk.END)
            self.log_text.config(state=tk.DISABLED)

        self.root.after(100, self.update_logs)  # Check queue periodically

    def download_mp3(self, url, output_path):
        """Download an MP3 file."""
        try:
            response = requests.get(url, stream=True, timeout=10)
            response.raise_for_status()
            with open(output_path, "wb") as f:
                for chunk in response.iter_content(chunk_size=8192):
                    f.write(chunk)
            return True
        except Exception as e:
            return str(e)

    def toggle_pause(self):
        """Pause or resume the download process."""
        if not self.is_downloading:
            return

        self.is_paused = not self.is_paused
        if self.is_paused:
            self.pause_button.config(text="Resume")
            self.log_message("Paused...")
        else:
            self.pause_button.config(text="Pause")
            self.log_message("Resumed...")

    def stop_download(self):
        """Stop the download process."""
        if not self.is_downloading:
            return

        self.stop_thread = True
        self.is_paused = False
        self.pause_button.config(state=tk.DISABLED)
        self.stop_button.config(state=tk.DISABLED)
        self.log_message("Stopping download...")

    def start_download_thread(self):
        """Start the download process in a separate thread."""
        if not self.excel_file_path:
            messagebox.showerror("Error", "No Excel file selected!")
            return

        self.start_button.config(state=tk.DISABLED)
        self.pause_button.config(state=tk.NORMAL)
        self.stop_button.config(state=tk.NORMAL)
        self.stop_thread = False

        threading.Thread(target=self.start_download, daemon=True).start()

    def start_download(self):
        """Perform the download process."""
        self.is_downloading = True
        self.log_message("Starting download process...")

        try:
            df = pd.read_excel(self.excel_file_path)
        except Exception as e:
            self.log_message(f"Failed to read Excel file: {e}")
            self.reset_buttons()
            return

        if "فایل" not in df.columns:
            self.log_message("Excel file must contain a 'فایل' column!")
            self.reset_buttons()
            return

        if Path(self.download_log_path).exists():
            downloaded_log = pd.read_csv(self.download_log_path)
        else:
            downloaded_log = pd.DataFrame(columns=["row_id", "mp3_link", "status"])

        downloaded_rows = set(downloaded_log["row_id"])

        for index, row in df.iterrows():
            if self.stop_thread:
                self.log_message("Download stopped.")
                break

            row_id = index + 1
            url = row["فایل"]
            ssn = row["کد ملی"]
            call_date = row["تاریخ شروع تماس"]

            if pd.isna(url):
                self.log_message(f"Row {row_id}: Empty.")
                continue

            if row_id in downloaded_rows:
                self.log_message(f"Row {row_id}: Already downloaded.")
                continue

            while self.is_paused:
                if self.stop_thread:
                    self.log_message("Download stopped during pause.")
                    self.reset_buttons()
                    return

            filename = f"{row_id}_{ssn}_{call_date}_{os.path.basename(url)}"
            output_path = os.path.join(self.download_directory, filename)
            os.makedirs(self.download_directory, exist_ok=True)

            self.log_message(f"Row {row_id}: Downloading...")
            result = self.download_mp3(url, output_path)

            if result is True:
                self.log_message(f"Row {row_id}: Downloaded successfully.")
                new_entry = {"row_id": row_id, "mp3_link": url, "status": "Success"}
            else:
                self.log_message(f"Row {row_id}: Failed to download ({result}).")
                new_entry = {"row_id": row_id, "mp3_link": url, "status": f"Failed: {result}"}

            # Update and save the log
            downloaded_log = pd.concat([downloaded_log, pd.DataFrame([new_entry])], ignore_index=True)
            downloaded_log.to_csv(self.download_log_path, index=False)

        self.reset_buttons()
        self.log_message("Download process completed!")

    def reset_buttons(self):
        """Reset buttons and states after download."""
        self.root.after(0, self._reset_buttons_safe)

    def _reset_buttons_safe(self):
        """Safely reset buttons."""
        self.is_downloading = False
        self.pause_button.config(state=tk.DISABLED)
        self.stop_button.config(state=tk.DISABLED)
        self.start_button.config(state=tk.NORMAL)

    def on_closing(self):
        """Handle application close event."""
        self.stop_thread = True
        self.root.destroy()


if __name__ == "__main__":
    root = tk.Tk()
    app = MP3DownloaderApp(root)
    root.mainloop()
