# main.py - Auto File Renamer (Final Pro Version)
import customtkinter as ctk
import tkinter as tk
from tkinter import filedialog, messagebox
import os
import shutil
from pathlib import Path
import csv
import json
import logging
from datetime import datetime
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler
import re
import webbrowser
import requests
import sys
from PIL import Image, ImageTk
import winsound
import warnings
warnings.filterwarnings("ignore", category=UserWarning, module="customtkinter")
warnings.filterwarnings("ignore", category=UserWarning, module="winsound")

# ============ Resolve Project Directory ============
if getattr(sys, 'frozen', False):
    project_dir = Path(sys.executable).parent
else:
    project_dir = Path(__file__).parent
    SOUND_FILE = project_dir / "sounds" / "ding.wav"


def resource_path(relative_path):
    """Get absolute path to resource, works for dev and PyInstaller"""
    try:
        base_path = sys._MEIPASS  # PyInstaller creates a temp folder and stores path in _MEIPASS
    except Exception:
        base_path = project_dir
    return Path(base_path) / relative_path

# ============ Create Directories ============
logs_dir = project_dir / "logs"
codes_dir = project_dir / "codes"
config_dir = project_dir / "config"
backup_dir = project_dir / "backup"
logs_dir.mkdir(exist_ok=True)
codes_dir.mkdir(exist_ok=True)
config_dir.mkdir(exist_ok=True)
backup_dir.mkdir(exist_ok=True)

# ============ Keyword Config Path ============
keywords_file = config_dir / "keywords.json"
DEFAULT_KEYWORDS = ["copy", "copies", "pcs", "pieces", "x"]

# ============ Configure Logging ============
log_file = logs_dir / f"{datetime.now().strftime('%Y-%m-%d')}.log"
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s | %(levelname)s | %(message)s",
    handlers=[
        logging.FileHandler(log_file, encoding="utf-8"),
        logging.StreamHandler()
    ]
)

# ============ Set Appearance ============
ctk.set_appearance_mode("Dark")
ctk.set_default_color_theme("blue")

# ============ Fonts ============
FONT = ("Segoe UI", 12)
TITLE_FONT = ("Segoe UI", 16, "bold")
CODE_FONT = ("Consolas", 13)
SMALL_FONT = ("Segoe UI", 10)

# ============ Version Info ============
APP_NAME = "Auto File Renamer Pro"
VERSION = "2.0.0"
AUTHOR = "Jignesh Thummar"
COPYRIGHT = "© 2025 Jignesh Thummar. All Rights Reserved. V2.0.0"
LICENSE = "Proprietary Software. Do not distribute."

class RenameHistory:
    def __init__(self):
        self.history = []
        self.index = -1

    def add(self, old, new):
        self.history = self.history[:self.index + 1]
        self.history.append({
            "old": str(old),
            "new": str(new),
            "timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        })
        self.index += 1

    def undo(self):
        if self.index < 0:
            return None
        item = self.history[self.index]
        self.index -= 1
        return item

    def redo(self):
        if self.index >= len(self.history) - 1:
            return None
        self.index += 1
        return self.history[self.index]

    def clear(self):
        self.history = []
        self.index = -1

    def count_renamed_in_done(self, root):
        done_count = 0
        for item in self.history:  # ✅ Correct — matches your actual variable
            if "Done" in str(item["new"]) or str(root) in str(item["new"]):
                done_count += 1
        return done_count


class FileRenamerApp(ctk.CTk):
    class AutoScanHandler(FileSystemEventHandler):
        def __init__(self, app):
            self.app = app
        def on_created(self, event):
            if not event.is_directory:
                ext = Path(event.src_path).suffix.lower()
                if ext in self.app.allowed_extensions:
                    self.app.after(0, self.app.scan_folder)    
    def __init__(self):
        super().__init__()
        # Set window and taskbar icon
        try:
            icon_path = project_dir / "icon.ico"
            if icon_path.exists():
                self.iconbitmap(icon_path)
            else:
                print("Icon file not found:", icon_path)
        except Exception as e:
            print("Icon error:", e)
        self.title(f"{APP_NAME} v{VERSION}")
        self.geometry("1000x780")
        self.resizable(True, True)

        # ============ Version & Auto-Update ============
        self.CURRENT_VERSION = "1.6.0"
        self.UPDATE_URL = "https://raw.githubusercontent.com/jigsthummar007/AutoFileRenamer-Pro/main/version.txt"
        
        # --- Paths ---
        self.config_file = project_dir / "config.json"

        # --- Data ---
        self.selected_root = None
        self.selected_file = None
        self.allowed_extensions = {'.plt', '.jpg', '.jpeg', '.jpe', '.jfif'}
        self.party_map = {}
        self.history = RenameHistory()
        self.auto_observer = None
        self.file_path_list = []
        self.filtered_file_list = []  # For search
        self.machine_var = ctk.StringVar(value="(C.S)")
        self.show_done_var = ctk.BooleanVar(value=False)
        self.last_folder = ""
        self.quantity_keywords = []
        self.first_run = True

        # --- Load Config ---
        self.load_config()
        self.load_keywords()



        # --- Menu Bar ---
        self.create_menu_bar()
        # Window close handler
        self.protocol("WM_DELETE_WINDOW", self.on_closing)

        # ============ Grid Layout ============
        # Configure grid layout
        # Configure rows
        self.grid_columnconfigure(1, weight=1)
        for i in range(10):
            self.grid_rowconfigure(i, weight=0)
        self.grid_rowconfigure(4, weight=1)  # Only file list grows

        # ============ Sidebar ============
        self.sidebar_frame = ctk.CTkFrame(self, width=260, corner_radius=12, fg_color="#2a2d30")
        self.sidebar_frame.grid(row=0, column=0, rowspan=8, sticky="nswe", padx=14, pady=14)  # Increased padding
        self.sidebar_frame.grid_propagate(False)

        # Add internal padding at top
        ctk.CTkLabel(self.sidebar_frame, text="").pack(pady=(0, 2))  # Top spacer        
        ctk.CTkLabel(self.sidebar_frame, text="📁 File Manager", font=TITLE_FONT).pack(pady=(12, 8))

        # --- Logo Image ---
        try:
            from PIL import Image
            logo_path = project_dir / "logo.png"
            if logo_path.exists():
                logo_image = ctk.CTkImage(
                    light_image=Image.open(logo_path),
                    dark_image=Image.open(logo_path),
                    size=(130, 60)  # Width, Height
                )
                self.logo_label = ctk.CTkLabel(
                    self.sidebar_frame,
                    image=logo_image,
                    text=""
                )
                self.logo_label.pack(pady=(0, 10))
            else:
                # Fallback: Text label if logo not found
                ctk.CTkLabel(
                    self.sidebar_frame,
                    text="Auto File Renamer Pro",
                    font=("Segoe UI", 11, "italic"),
                    text_color="gray70"
                ).pack(pady=(0, 10))
        except Exception as e:
            # If PIL or file error, skip logo
            pass

        self.folder_btn = ctk.CTkButton(
            self.sidebar_frame,
            text="📁 Select Folder",
            command=self.select_folder,
            height=40,
            font=FONT
        )
        self.folder_btn.pack(pady=10, padx=16, fill="x")

        self.scan_btn = ctk.CTkButton(
            self.sidebar_frame,
            text="⟳ Scan Files",
            command=self.scan_folder,
            height=35
        )
        self.scan_btn.pack(pady=8, padx=16, fill="x")

        # --- Auto-scan ---
        self.auto_scan_var = ctk.BooleanVar(value=True)
        self.auto_scan_switch = ctk.CTkSwitch(
            self.sidebar_frame,
            text="🔁 Auto-scan",
            variable=self.auto_scan_var,
            command=self.toggle_auto_scan,
            font=FONT
        )
        self.auto_scan_switch.pack(pady=12, padx=16, anchor="w")

        # --- Show Done Files ---
        self.show_done_switch = ctk.CTkSwitch(
            self.sidebar_frame,
            text="👁️ Show Finalize Files",
            variable=self.show_done_var,
            command=self.on_finalize_mode_change,
            font=FONT
        )
        self.show_done_switch.pack(pady=10, padx=16, anchor="w")

        # --- Machine Type ---
        ctk.CTkLabel(self.sidebar_frame, text="🖨️ Machine:", font=FONT).pack(pady=(14, 4), anchor="w", padx=20)
        self.machine_dropdown = ctk.CTkComboBox(
            self.sidebar_frame,
            values=["(C.S)", "(C.E)"],
            variable=self.machine_var,
            font=FONT,
            dropdown_font=FONT
        )
        self.machine_dropdown.pack(pady=8, padx=20, fill="x")

        # --- Edit Keywords Button ---
        self.keywords_btn = ctk.CTkButton(
            self.sidebar_frame,
            text="🔧 Edit Quantity Keywords",
            command=self.open_keywords_editor,
            fg_color="orange",
            hover_color="dark orange",
            height=32
        )
        self.keywords_btn.pack(pady=10, padx=16, fill="x")

        # --- Reload CSV ---
        self.reload_csv_btn = ctk.CTkButton(
            self.sidebar_frame,
            text="🔁 Reload Parties",
            command=self.load_parties_csv,
            fg_color="#8a2be2",
            hover_color="#7a1dd1",
            height=32
        )
        self.reload_csv_btn.pack(pady=14, padx=16, fill="x")

        # --- Manage Parties Button ---
        self.manage_parties_btn = ctk.CTkButton(
            self.sidebar_frame,
            text="👥 Manage Parties",
            command=self.open_parties_editor,
            fg_color="purple",
            hover_color="#4b0082",
            height=32
        )
        self.manage_parties_btn.pack(pady=10, padx=16, fill="x")

        # --- Export Log Button ---
        self.export_log_btn = ctk.CTkButton(
            self.sidebar_frame,
            text="📊 Export Rename Log",
            command=self.export_rename_log,
            fg_color="teal",
            hover_color="#006666",
            height=32
        )
        self.export_log_btn.pack(pady=10, padx=16, fill="x")

        # ============ Main Area ============
        self.main_frame = ctk.CTkFrame(self, corner_radius=12, fg_color="transparent")  # Optional: match background
        self.main_frame.grid(row=0, column=1, sticky="nswe", padx=14, pady=14)  # Match sidebar
        self.main_frame.grid_columnconfigure(0, weight=1)
        self.main_frame.grid_rowconfigure(3, weight=1)

        # --- Search Box ---
        search_frame = ctk.CTkFrame(self.main_frame, fg_color="transparent")
        search_frame.grid(row=0, column=0, columnspan=2, sticky="we", pady=(0, 5))
        ctk.CTkLabel(search_frame, text="🔍 Search:", font=FONT).pack(side="left")
        self.search_var = ctk.StringVar()
        self.search_var.trace("w", self.on_search_change)
        self.search_entry = ctk.CTkEntry(search_frame, textvariable=self.search_var, placeholder_text="Filter files...")
        self.search_entry.pack(side="left", fill="x", expand=True, padx=(5, 0))

        # - Selected File -
        self.file_label = ctk.CTkLabel(
            self.main_frame,
            text="🎯 Select a file to preview rename",
            font=("Segoe UI", 16, "bold"),
            text_color="grey70",
            anchor="w"
        )
        self.file_label.grid(row=1, column=0, columnspan=2, sticky="w", padx=10, pady=(10, 0))

        # - Folder Path Label (Normal Mode Only) -
        self.folder_path_label = ctk.CTkLabel(
            self.main_frame,
            text="📍 Folder: —",
            font=("Segoe UI", 12),
            text_color="#FF7E33",
            anchor="w"
        )
        self.folder_path_label.grid(row=2, column=0, columnspan=2, sticky="w", padx=10, pady=(6, 0))

        # - Preview -
        self.preview_label = ctk.CTkLabel(
            self.main_frame,
            text="Preview: --",
            font=("Consolas", 14),
            text_color="#aaffaa",
            anchor="w"
        )
        self.preview_label.grid(row=3, column=0, columnspan=2, sticky="w", padx=10, pady=(6, 8))

        # - Buttons -
        btn_frame = ctk.CTkFrame(self.main_frame, fg_color="transparent")
        btn_frame.grid(row=4, column=0, columnspan=2, sticky="w", pady=(12, 12))

        # Test Button
        self.test_btn = ctk.CTkButton(
            btn_frame,
            text="🔍 Test",
            command=self.test_rename,
            fg_color="#2734F5",
            hover_color="#000F66",
            width=80,
            height=32,
            font=("Segoe UI", 12),
        )
        self.test_btn.grid(row=0, column=0, padx=(0, 6))

        # Undo Button
        self.undo_btn = ctk.CTkButton(
            btn_frame,
            text="↩ Undo",
            command=self.undo_rename,
            width=80,
            height=32,
            fg_color="red",
            hover_color="#8B0000",
            font=("Segoe UI", 12)
        )
        self.undo_btn.grid(row=0, column=1, padx=(0, 6))

        # Redo Button
        self.redo_btn = ctk.CTkButton(
            btn_frame,
            text="⟳ Redo",
            command=self.redo_rename,
            fg_color="#2734F5",
            hover_color="#000F66",
            width=80,
            height=32,
            font=("Segoe UI", 12)
        )
        self.redo_btn.grid(row=0, column=2, padx=(0, 12))

        # Rename Button
        self.rename_btn = ctk.CTkButton(
            btn_frame,
            text="✅ Rename",
            command=self.rename_file,
            fg_color="#008800",
            hover_color="#006600",
            width=100,
            height=32,
            font=("Segoe UI", 12, "bold")
        )
        self.rename_btn.grid(row=0, column=3)

        self.select_all_btn = ctk.CTkButton(
            btn_frame,
            text="📁 Batch Rename",
            command=self.select_all_files,
            fg_color="#008800",
            hover_color="#006600",
            width=120,           # Match Batch Rename width
            height=32,           # Match top button height
            font=("Segoe UI", 14)
        )
        self.select_all_btn.grid(row=0, column=4, padx=(12, 6))

        self.undo_all_btn = ctk.CTkButton(
            btn_frame,
            text="↩ Undo All",
            command=self.undo_all_batch,
            width=120,
            height=32,
            fg_color="red",
            hover_color="#8B0000",
            font=("Segoe UI", 14)
        )
        self.undo_all_btn.grid(row=0, column=5, padx=(0, 0))

        # - File List: Use tk.Text for tag support -
        self.file_listbox = tk.Text(
            self.main_frame,
            font=("Consolas", 12),
            wrap="none",
            bg="#2a2d30",           # Dark background
            fg="lightgray",         # Light gray text
            insertbackground="white",
            selectbackground="#1f538d",
            selectforeground="white",
            relief="flat",
            padx=10,
            pady=10,
            borderwidth=1,
            highlightthickness=1,
            highlightbackground="gray30"
        )
        self.file_listbox.grid(row=5, column=0, columnspan=2, sticky="nswe", pady=(6, 8))
        self.file_listbox.bind("<ButtonRelease-1>", self.on_file_click)
        self.file_listbox.configure(cursor="hand2")

        # Configure tags for styling
        self.file_listbox.tag_configure("header", foreground="cyan", font=("Consolas", 12, "bold"))
        self.file_listbox.tag_configure("selected", background="#1f538d", foreground="white")
        self.file_listbox.tag_configure("normal", foreground="lightgray")

        
     
        # - Auto-Scrolling Bulletin (One Tip at a Time) -
        self.bulletin_messages = [
            "💡 Tip: Press Ctrl+A to batch rename all files",
            "💡 Tip: Click any file to preview rename",
            "💡 Tip: Edit keywords for 'x2', 'copy', 'pcs'",
            "✅ Tip: Dont Forget to Chnage Eco & Solvant Mode Before Rename",
            "💡 Tip: You can mage Party and Code. Click Manage Party Button",
            "✅ Renamed files go to the 'Done' folder",
            "🔔 Press F1 for help anytime!"
        ]

        # Add Easter Egg messages
        self.bulletin_messages.extend([
            "🎉 Wow! All files renamed! You're on fire!",
            "🚀 Auto File Renamer Pro loves you!",
            "🔥 You just saved 2 hours of manual work!",
            "💡 Pro Tip: You're awesome at this!",
            "🏆 Achievement Unlocked: Batch Master!"
        ])

        self.current_msg_index = 0
        self.bulletin_offset = 0
        self.is_scrolling = True

        # - Bulletin Container (Force Left Alignment) -
        bulletin_container = ctk.CTkFrame(
            self.main_frame,
            fg_color="gray20",
            corner_radius=6,
            height=30
        )
        bulletin_container.grid(row=6, column=0, columnspan=2, sticky="we", padx=10, pady=(6, 0))
        bulletin_container.grid_propagate(False)
        bulletin_container.grid_rowconfigure(0, weight=1)
        bulletin_container.grid_columnconfigure(0, weight=1)

        # - Bulletin Label (Left-aligned via pack) -
        self.bulletin_label = ctk.CTkLabel(
            bulletin_container,
            text="",
            font=("Segoe UI", 14, "bold"),
            text_color="lightblue",
            fg_color="transparent",
            anchor="w"
        )
        self.bulletin_label.pack(side="left", padx=0, pady=0, anchor="w")

        # Start scrolling
        self.after(500, self.scroll_bulletin)

        # --- Status Bar ---
        self.status_label = ctk.CTkLabel(
            self,
            text="Ready",
            anchor="w",
            font=("Segoe UI", 14),
            text_color="lightgray"
        )
        self.status_label.grid(row=8, column=0, columnspan=2, sticky="we", padx=18, pady=5)

        # --- Math Input (Below Bulletin, Same Width) ---
        math_frame = ctk.CTkFrame(self.main_frame, fg_color="transparent", height=32)
        math_frame.grid(row=7, column=0, columnspan=2, sticky="we", padx=10, pady=(6, 8))
        math_frame.grid_propagate(False)

        # Icon + Entry
        ctk.CTkLabel(math_frame, text="🧮", font=("Segoe UI", 13)).pack(side="left", padx=(0, 8))

        self.math_entry = ctk.CTkEntry(
            math_frame,
            placeholder_text="Math: 30*2+10",
            height=28,
            font=("Segoe UI", 11),
            width=300  # Adjust if needed
        )
        self.math_entry.pack(side="left", fill="x", expand=True)

        # Bind Enter
        self.math_entry.bind("<Return>", self.calculate_math)

        # - Redesigned Footer (Row 9) -
        footer_frame = ctk.CTkFrame(self, fg_color="transparent")
        footer_frame.grid(row=9, column=0, columnspan=2, pady=(0, 6), sticky="we", padx=18)

        # Use grid inside footer_frame for precise control
        footer_frame.grid_columnconfigure(0, weight=1)  # Left
        footer_frame.grid_columnconfigure(1, weight=2)  # Center
        footer_frame.grid_columnconfigure(2, weight=1)  # Right

        # --- Left: Info Label ---
        self.info_label = ctk.CTkLabel(
            footer_frame,
            text="📁 Files: 0 | ✅ Renamed: 0 | ⏳ Pending: 0 | 🕒 Last: --:-- | 📄 Parties: 0",
            font=("Segoe UI", 13, "bold"),
            text_color="#ffffff",
            anchor="w"
        )
        self.info_label.grid(row=0, column=0, sticky="w", padx=(0, 10))

        # --- Center: Copyright ---
        self.copyright_label = ctk.CTkLabel(
            footer_frame,
            text=COPYRIGHT,
            font=("Segoe UI", 13),
            text_color="#ff5f5f"
        )
        self.copyright_label.grid(row=0, column=1, sticky="nsew")

        # --- Right: Contact Icons ---
        contact_frame = ctk.CTkFrame(footer_frame, fg_color="transparent")
        contact_frame.grid(row=0, column=2, sticky="e")

        # WhatsApp
        ctk.CTkButton(
            contact_frame,
            text="📞",
            width=36,
            height=28,
            font=("Segoe UI", 14),
            fg_color="gray40",
            hover_color="gray50",
            command=lambda: webbrowser.open("https://wa.me/919825531314")
        ).pack(side="left", padx=(0, 6))

        # Instagram
        ctk.CTkButton(
            contact_frame,
            text="📸",
            width=36,
            height=28,
            font=("Segoe UI", 14),
            fg_color="gray40",
            hover_color="gray50",
            command=lambda: webbrowser.open("https://instagram.com/official.jignesh.1")
        ).pack(side="left", padx=(0, 6))

        # Email
        ctk.CTkButton(
            contact_frame,
            text="📧",
            width=36,
            height=28,
            font=("Segoe UI", 14),
            fg_color="gray40",
            hover_color="gray50",
            command=lambda: webbrowser.open("mailto:Jigsthummar1990@gmail.com")
        ).pack(side="left")

        # ============ Keyboard Shortcuts ============
        self.bind("<Control-o>", lambda e: self.select_folder())
        self.bind("<Control-s>", lambda e: self.scan_folder())
        self.bind("<Control-r>", lambda e: self.rename_file() if self.selected_file else None)
        self.bind("<Control-z>", lambda e: self.undo_rename())
        self.bind("<Control-y>", lambda e: self.redo_rename())
        self.bind("<Control-a>", lambda e: self.select_all_files())

        # Hide main window until splash is gone
        self.withdraw()

        # Show splash
        self.show_splash_screen()

        # Start auto-scan and load data
        self.toggle_auto_scan()

        # --- Load last folder ---
        if self.last_folder and Path(self.last_folder).exists():
            self.selected_root = Path(self.last_folder)
            self.scan_folder()
            self.start_auto_scan()

        self.load_parties_csv()
        self.update_info_bar()

        # --- Create Backup (must come after status_label is created)
        self.create_backup()
        self.update_info_bar()
        
        # Initial button visibility
        self.update_button_visibility()

    def test_rename(self):
        """Show what the selected file would be renamed to — without doing anything"""
        if not self.selected_file or not self.selected_file.exists():
            self.status_label.configure(text="❌ No file selected")
            return

        file_path = self.selected_file
        party_folder = self.find_party_folder(file_path)
        if not party_folder:
            self.status_label.configure(text="❌ No party folder found")
            return

        # Find original party name (case-sensitive from CSV)
        original_party_name = next(
            (name for name in self.party_map if name.lower() == party_folder.name.lower()),
            party_folder.name
        )
        party_code = self.party_map.get(original_party_name, "?")

        # Get parts
        stem = file_path.stem
        ext = file_path.suffix
        dim_str = self.extract_dimensions(stem)
        new_name = self.generate_new_filename(stem, party_code, ext, dim_str)

        # Show result
        msg = (
            f"📄 Selected File:\n{file_path.name}\n\n"
            f"🔢 Would Be Renamed To:\n{new_name}\n\n"
            f"🏷️ Party: {original_party_name} ({party_code})\n"
            f"🖨️ Machine: {self.machine_var.get()}\n"
            f"📐 Dimensions: {dim_str}\n"
            f"📦 Quantity: {self.detect_quantity(stem)}\n\n"
            f"✅ This is just a preview. No file was changed."
        )
        messagebox.showinfo("🔍 Test Mode: Rename Preview", msg)
        self.status_label.configure(text=f"🔍 Previewed: {file_path.name}")

    def update_folder_path_display(self):
        """Update folder path label with clean formatting — only in Normal Mode"""
        if not self.show_done_var.get() and self.selected_root:
            # Get parts of the path
            parts = self.selected_root.parts
            # Show last 4 parts
            display_parts = []
            for part in parts[-4:]:
                # Clean up common issues
                if part.lower() == "05 may":
                    part = "May"
                elif part.lower() == "01-6":
                    part = "1–6"
                elif part.lower() == "gatam cmrk":
                    part = "Gatam CMYK"
                else:
                    part = part.replace("_", " ")  # Replace underscores
                display_parts.append(part)
            display_text = " > ".join(display_parts)
            self.folder_path_label.configure(text=f"📍 Folder: {display_text}")
            self.folder_path_label.grid()  # Ensure visible
        else:
            self.folder_path_label.grid_remove()  # Hide in Finalize Mode

    def update_info_bar(self):
        """Update the persistent info bar with current stats and force color"""
        total = len(self.file_path_list)
        renamed = self.history.count_renamed_in_done(self.selected_root) if self.selected_root else 0
        pending = total - renamed
        last_time = datetime.now().strftime("%H:%M")
        num_parties = len(self.party_map)

        # Count finalized files (with [ok])
        finalized = 0
        if self.selected_root:
            for file_path in self.selected_root.rglob("*[ok]*"):
                if "[ok]" in file_path.name:
                    finalized += 1

        machine = self.machine_var.get()
        auto_status = "ON" if self.auto_scan_var.get() else "OFF"

        info_text = (
            f"📁 Files: {total} | "
            f"✅ Renamed: {renamed} | "
            f"⏳ Pending: {pending} | "
            f"✔ Finalized: {finalized} | "
            f"🖨️ {machine} | "
            f"🔁 Auto: {auto_status} | "
            f"🕒 {last_time} | "
            f"📄 Parties: {num_parties}"
        )

        # --- Force Strong Colors ---
        if pending == 0 and total > 0:
            text_color = "#4cd964"  # Apple-style green
        elif pending > 0:
            text_color = "#ffcc00"  # Soft yellow (not white)
        else:
            text_color = "#8e8e93"  # iOS-style gray

        # ✅ Force update with text_color
        self.info_label.configure(
            text=info_text,
            text_color=text_color  # This will now work
        )

    def show_splash_screen(self):
        """Show animated GIF splash screen at startup"""
        splash = ctk.CTkToplevel(self)
        splash.overrideredirect(True)
        splash.configure(fg_color="white")
    
        # 🔧 Define image_path here
        image_path = resource_path("splash.gif")
    
        try:
            from PIL import Image, ImageTk
        
            # Check if file exists
            if not image_path.exists():
                raise FileNotFoundError(f"Splash image not found: {image_path}")
        
            # Load and prepare GIF
            gif = Image.open(image_path)
            frames = []
        
            # Extract all frames
            try:
                frame_index = 0
                while True:
                    # Convert frame to RGB (or RGBA) to support transparency/colors
                    frame = gif.copy().convert("RGBA")
                    # Resize for display
                    resized_frame = frame.resize((500, 300), Image.Resampling.LANCZOS)
                    # Convert to Tkinter-compatible PhotoImage
                    photo_image = ImageTk.PhotoImage(resized_frame)
                    frames.append(photo_image)
                
                    frame_index += 1
                    gif.seek(frame_index)  # Go to next frame
                
            except EOFError:
                pass  # End of frames reached
        
            if not frames:
                raise Exception("No frames loaded from GIF")
        
            # Create label to hold image
            img_label = ctk.CTkLabel(splash, text="")
            img_label.pack(expand=True)
        
            # Animation loop
            def animate_gif(frame_idx=0):
                img_label.configure(image=frames[frame_idx])
                splash.after(100, animate_gif, (frame_idx + 1) % len(frames))  # Loop continuously
        
            animate_gif()
        
        except Exception as e:
            # Fallback: Show text if GIF fails
            ctk.CTkLabel(splash, text="Auto File Renamer", font=("Helvetica", 24, "bold")).pack(pady=40)
            ctk.CTkLabel(splash, text="Loading...", font=("Helvetica", 16)).pack()

        # Center splash
        splash.update_idletasks()
        x = (self.winfo_screenwidth() // 2) - (500 // 2)
        y = (self.winfo_screenheight() // 2) - (300 // 2)
        splash.geometry(f"500x300+{int(x)}+{int(y)}")

        # Store reference and schedule close
        self.splash = splash
        self.after(4000, self.destroy_splash)  # Show splash for 4 seconds

    def destroy_splash(self):
        """Close splash and show main app"""
        if hasattr(self, 'splash'):
            self.splash.destroy()

        try:
            # Now safe to load data and start
            self.load_config()
            self.load_keywords()
            self.load_parties_csv()  # May crash if CSV is missing
            self.toggle_auto_scan()

            if self.last_folder and Path(self.last_folder).exists():
                self.selected_root = Path(self.last_folder)
                self.scan_folder()
                self.start_auto_scan()

            self.update_info_bar()
            self.create_backup()
            self.check_for_update()  # Optional: check after startup

            # ✅ Must be last: show window
            self.deiconify()  # Show main window
            self.lift()
            self.focus_force()

        except Exception as e:
            # ✅ Show error even if window was hidden
            import tkinter.messagebox as mbox
            mbox.showerror("Startup Error", f"Failed to start:\n{e}")
            self.destroy()  # Close app

    def on_finalize_mode_change(self):
        """Called when 'Show Finalize Files' is toggled"""
        self.scan_folder()
        self.update_button_visibility()  
        self.update_folder_path_display()   

    def create_menu_bar(self):
        """Create menu bar with Help and About"""
        menubar = tk.Menu(self)
        
        # Help Menu
        help_menu = tk.Menu(menubar, tearoff=0)
        help_menu.add_command(label="Usage Guide", command=self.show_help_usage)
        help_menu.add_command(label="Keyboard Shortcuts", command=self.show_help_shortcuts)
        menubar.add_cascade(label="📘 Help", menu=help_menu)

        # About Menu
        menubar.add_command(label="ℹ️ About", command=self.show_about)

        self.configure(menu=menubar)

    def show_help_shortcuts(self):
        """Show keyboard shortcuts"""
        msg = """
📌 Keyboard Shortcuts:

• Ctrl + O → Select Folder
• Ctrl + S → Scan Files
• Ctrl + R → Rename Selected File
• Ctrl + Z → Undo Last Rename
• Ctrl + Y → Redo Rename
• Ctrl + A → Batch Rename All Files

💡 Tip: Click any file to preview rename.
💡 Tip: Use 'Show Done Files' to finalize with %8%.
"""
        messagebox.showinfo("📘 Help: Keyboard Shortcuts", msg)

    def show_help_usage(self):
        """Show how to use the app"""
        msg = """
📘 How to Use Auto File Renamer:

1. Click '📁 Select 2025 Folder' to set root
2. Files will appear in the list
3. Click any file to see rename preview
4. Click '✅ Rename' to process
5. File moves to 'Done' folder automatically
6. Use '📁 Batch Rename' for multiple files
7. Use '🔧 Edit Keywords' to add 'layout', 'design', etc.
8. Use '👁️ Show Finalize Files' to manually add %8%

📁 Folder Structure:
2025 → Month → Date → Party → File.plt

📤 Output:
{Code}_{Name} (C.S)(FT.1x2)(Q.2)%%.plt → Moved to 'Done'

🔧 Keywords:
- Use 'x' → '2 x' → (Q.2)
- Add more via UI
"""
        messagebox.showinfo("📘 Help: Usage Guide", msg)

    def show_about(self):
        """Show about window"""
        popup = ctk.CTkToplevel(self)
        popup.title("ℹ️ About")
        popup.geometry("450x500")
        popup.resizable(False, False)
        popup.transient(self)
        popup.grab_set()

        x = self.winfo_x() + (self.winfo_width() // 2) - 225
        y = self.winfo_y() + (self.winfo_height() // 2) - 250
        popup.geometry(f"+{int(x)}+{int(y)}")

        ctk.CTkLabel(popup, text=APP_NAME, font=("Segoe UI", 18, "bold")).pack(pady=10)
        ctk.CTkLabel(popup, text=f"Version {VERSION}", font=("Segoe UI", 14)).pack(pady=5)
        ctk.CTkLabel(popup, text=f"By: {AUTHOR}", font=("Segoe UI", 14)).pack(pady=5)

        ctk.CTkLabel(popup, text="📱 Contact:", font=("Segoe UI", 12, "bold")).pack(pady=(20, 5))
        ctk.CTkLabel(popup, text="WhatsApp: +91 98255 31314", font=("Segoe UI", 12)).pack()
        ctk.CTkLabel(popup, text="Instagram: @official.jignesh.1", font=("Segoe UI", 12)).pack()
        ctk.CTkLabel(popup, text="Email: Jigsthummar1990@gmail.com", font=("Segoe UI", 12)).pack()

        ctk.CTkLabel(popup, text="📜 License", font=("Segoe UI", 12, "bold")).pack(pady=(20, 5))
        ctk.CTkLabel(popup, text=COPYRIGHT, font=("Segoe UI", 10), wraplength=400).pack(pady=5)
        ctk.CTkLabel(popup, text="Proprietary Software. Do not distribute.", font=("Segoe UI", 10), wraplength=400).pack(pady=5)

        ctk.CTkButton(
            popup, text="📧 Send Email", width=120,
            command=lambda: webbrowser.open("mailto:Jigsthummar1990@gmail.com")
        ).pack(pady=10)

        ctk.CTkButton(popup, text="Close", width=100, command=popup.destroy).pack(pady=10)

    def update_button_visibility(self):
        """Show or hide buttons based on finalize mode"""
        if self.show_done_var.get():
            # Finalize Mode: Hide Rename, Undo, Redo, Undo All
            self.rename_btn.grid_remove()
            self.undo_btn.grid_remove()
            self.redo_btn.grid_remove()
            self.undo_all_btn.grid_remove()
            self.select_all_btn.grid_remove()
        else:
            # Normal Mode: Show all buttons
            self.rename_btn.grid()
            self.undo_btn.grid()
            self.redo_btn.grid()
            self.undo_all_btn.grid()
            self.select_all_btn.grid()

    def calculate_math(self, event=None):
        """Evaluate simple math expression"""
        expr = self.math_entry.get().strip()
        if not expr:
            return
        try:
            # Allow only safe characters
            allowed = "0123456789+-*/(). %"
            if all(c in allowed for c in expr):
                result = eval(expr)
                if isinstance(result, float):
                    result = round(result, 2)
                self.status_label.configure(text=f"🧮 Math: {result} | Ready")
                self.math_entry.delete(0, "end")  # Clear after use
            else:
                self.status_label.configure(text="❌ Invalid characters in math")
        except:
            self.status_label.configure(text="❌ Math error")

    def scroll_bulletin(self):
        """Scroll bulletin text from LEFT to RIGHT, starting from the left edge"""
        if not self.is_scrolling or not self.bulletin_messages:
            return

        # Current message
        msg = self.bulletin_messages[self.current_msg_index]
        display_text = msg  # No extra spaces

        # Settings
        visible_width = 60  # Adjust based on your window width
        total_len = len(display_text)

        # Create a looped string
        looped = display_text + "   " + display_text  # Add gap

        # Start position moves left → visual right scroll
        start_pos = self.bulletin_offset % total_len

        # Get visible slice
        visible = looped[start_pos:start_pos + visible_width]

        # Update label
        self.bulletin_label.configure(text=visible)

        # Move to next position
        self.bulletin_offset += 1

        # Repeat
        self.after(150, self.scroll_bulletin)

    def run_startup_wizard(self):
        if not self.first_run:
            return
        self.first_run = False
        wizard = ctk.CTkToplevel(self)
        wizard.title("🚀 Setup Wizard")
        wizard.geometry("500x300")
        wizard.resizable(False, False)
        wizard.transient(self)
        wizard.grab_set()
        wizard.lift()
        wizard.focus_force()

        # Center popup
        x = self.winfo_x() + (self.winfo_width() // 2) - 250
        y = self.winfo_y() + (self.winfo_height() // 2) - 150
        wizard.geometry(f"+{int(x)}+{int(y)}")

        ctk.CTkLabel(wizard, text="Welcome to Auto File Renamer!", font=("Segoe UI", 16, "bold")).pack(pady=20)
        ctk.CTkLabel(
            wizard,
            text="This tool will help you rename files quickly with party codes, dimensions, and quantities.",
            wraplength=400,
            justify="center"
        ).pack(pady=10)
        ctk.CTkButton(wizard, text="Let's Go!", command=wizard.destroy).pack(pady=20)

        steps = [
            "1. Click '📁 Select 2025 Folder' to set root",
            "2. Your files will appear in the list",
            "3. Click any file to preview rename",
            "4. Use '🔧 Edit Keywords' to add 'layout', 'design', etc.",
            "5. Click '✅ Rename' to process",
            "6. Use '📁 Batch Rename' for multiple files"
        ]

        for step in steps:
            ctk.CTkLabel(wizard, text=step, font=("Segoe UI", 11), anchor="w").pack(pady=2, padx=20, anchor="w")

        def finish():
            self.first_run = False
            wizard.destroy()
            

        ctk.CTkButton(wizard, text="Let's Go!", command=finish, height=40, font=FONT).pack(pady=20)

    def check_for_update(self, force=False):
        """Check if a new version is available"""
        try:
            response = requests.get(self.UPDATE_URL, timeout=5)
            latest_version = response.text.strip()
            if latest_version > self.CURRENT_VERSION:
                self.show_update_prompt(latest_version)
            elif force:
                messagebox.showinfo("Check for Updates", f"You're on the latest version: {self.CURRENT_VERSION}")
        except Exception as e:
            if force:
                messagebox.showerror("Update Check Failed", f"Could not connect to update server:\n{e}")

    def show_update_prompt(self, latest_version):
        """Show popup when update is available"""
        popup = ctk.CTkToplevel(self)
        popup.title("🎉 Update Available")
        popup.geometry("400x150")
        popup.transient(self)
        popup.grab_set()

        ctk.CTkLabel(
            popup,
            text="A new version is available!",
            font=("Segoe UI", 16, "bold")
        ).pack(pady=10)

        ctk.CTkLabel(
            popup,
            text=f"Current: v{self.CURRENT_VERSION}\nLatest: v{latest_version}",
            font=("Consolas", 12)
        ).pack(pady=5)

        btn_frame = ctk.CTkFrame(popup, fg_color="transparent")
        btn_frame.pack(pady=10)

        def open_release_page():
            import webbrowser
            webbrowser.open("https://github.com/jigsthummar007/AutoRenamer/releases/latest")
            popup.destroy()

        ctk.CTkButton(btn_frame, text="Download Update", command=open_release_page).pack(side="left", padx=5)
        ctk.CTkButton(btn_frame, text="Close", command=popup.destroy).pack(side="left", padx=5)

    def create_backup(self):
        """Backup config, parties.csv, keywords.json"""
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        backup_folder = backup_dir / timestamp
        backup_folder.mkdir(exist_ok=True)

        items = [
            (self.config_file, "config.json"),
            (codes_dir / "parties.csv", "parties.csv"),
            (keywords_file, "keywords.json")
        ]

        for src, name in items:
            try:
                if src.exists():
                    shutil.copy2(src, backup_folder / name)
            except Exception as e:
                logging.warning(f"Backup failed for {name}: {e}")

        self.status_label.configure(text=f"✅ Backup created: {timestamp}")

    def export_rename_log(self):
        """Export rename history to CSV"""
        if not self.history.history:
            messagebox.showinfo("Export Log", "No rename history to export.")
            return

        file = filedialog.asksaveasfilename(
            defaultextension=".csv",
            filetypes=[("CSV files", "*.csv")],
            title="Save Rename Log"
        )
        if not file:
            return

        try:
            with open(file, mode='w', encoding='utf-8', newline='') as f:
                writer = csv.writer(f)
                writer.writerow(["Timestamp", "Party Folder", "Original", "New Name"])
                for item in self.history.history:
                    old_path = Path(item["old"])
                    party_folder = None
                    current = old_path.parent
                    while True:
                        if current.name in self.party_map:
                            party_folder = current.name
                            break
                        if str(current) == str(current.parent):
                            break
                        current = current.parent
                    party_folder = party_folder or "Unknown"
                    writer.writerow([item["timestamp"], party_folder, item["old"], item["new"]])
            messagebox.showinfo("Success", f"Rename log exported to:\n{file}")
        except Exception as e:
            messagebox.showerror("Export Failed", str(e))


    def load_keywords(self):
        """Load quantity keywords from config/keywords.json"""
        if keywords_file.exists():
            try:
                with open(keywords_file, "r", encoding="utf-8") as f:
                    data = json.load(f)
                    self.quantity_keywords = data.get("quantity_keywords", DEFAULT_KEYWORDS)
            except Exception as e:
                logging.warning(f"Failed to load keywords: {e}")
                self.quantity_keywords = DEFAULT_KEYWORDS
        else:
            self.quantity_keywords = DEFAULT_KEYWORDS
            self.save_keywords()

    def save_keywords(self):
        """Save keywords to file"""
        try:
            with open(keywords_file, "w", encoding="utf-8") as f:
                json.dump({"quantity_keywords": self.quantity_keywords}, f, indent=2)
        except Exception as e:
            logging.error(f"Failed to save keywords: {e}")

    def open_keywords_editor(self):
        """Open UI to edit quantity keywords"""
        popup = ctk.CTkToplevel(self)
        popup.title("🔧 Edit Quantity Keywords")
        popup.geometry("500x400")
        popup.resizable(False, False)
        popup.transient(self)
        popup.grab_set()

        x = self.winfo_x() + (self.winfo_width() // 2) - 250
        y = self.winfo_y() + (self.winfo_height() // 2) - 200
        popup.geometry(f"+{int(x)}+{int(y)}")

        ctk.CTkLabel(popup, text="Manage Quantity Keywords", font=TITLE_FONT).pack(pady=10)

        list_frame = ctk.CTkScrollableFrame(popup, height=200)
        list_frame.pack(pady=10, padx=20, fill="both", expand=True)

        keyword_vars = []
        delete_btns = []

        def refresh_list():
            for var, btn in zip(keyword_vars, delete_btns):
                var.destroy()
                btn.destroy()
            keyword_vars.clear()
            delete_btns.clear()

            for kw in self.quantity_keywords:
                inner_frame = ctk.CTkFrame(list_frame, fg_color="transparent")
                inner_frame.pack(fill="x", pady=2)

                var = ctk.CTkLabel(inner_frame, text=kw, font=CODE_FONT, width=100, anchor="w")
                var.pack(side="left", padx=(0, 10))
                keyword_vars.append(var)

                btn = ctk.CTkButton(
                    inner_frame,
                    text="🗑️",
                    width=40,
                    height=28,
                    fg_color="#d42222",
                    hover_color="#a00",
                    command=lambda k=kw: remove_keyword(k, popup)
                )
                btn.pack(side="right")
                delete_btns.append(btn)

        def remove_keyword(kw, win):
            self.quantity_keywords.remove(kw)
            refresh_list()

        def add_keyword():
            new_kw = add_entry.get().strip().lower()
            if new_kw and new_kw not in self.quantity_keywords:
                self.quantity_keywords.append(new_kw)
                add_entry.delete(0, "end")
                refresh_list()

        def save_and_close():
            self.save_keywords()
            popup.destroy()
            self.status_label.configure(text=f"✅ Keywords updated: {len(self.quantity_keywords)} active")

        def reset_default():
            self.quantity_keywords = DEFAULT_KEYWORDS.copy()
            refresh_list()

        refresh_list()

        add_frame = ctk.CTkFrame(popup)
        add_frame.pack(pady=10, padx=20, fill="x")
        ctk.CTkLabel(add_frame, text="Add New:").pack(side="left")
        add_entry = ctk.CTkEntry(add_frame, placeholder_text="e.g., layout, design")
        add_entry.pack(side="left", fill="x", expand=True, padx=5)
        ctk.CTkButton(add_frame, text="Add", command=add_keyword).pack(side="left")

        btn_frame = ctk.CTkFrame(popup, fg_color="transparent")
        btn_frame.pack(pady=10)
        ctk.CTkButton(btn_frame, text="💾 Save & Close", command=save_and_close).pack(side="left", padx=5)
        ctk.CTkButton(btn_frame, text="🔄 Reset", command=reset_default).pack(side="left", padx=5)

        popup.bind("<Return>", lambda e: add_keyword())
        popup.focus()

    def extract_dimensions(self, filename: str) -> str:
        """Extract dimensions and return FT string with custom height logic"""
        # Normalize: lowercase and clean extra spaces
        clean = re.sub(r'\s+', ' ', filename.strip().lower())
    
        # Convert all variations of 'x' and 'by' to a standard 'x' separator
        # This handles: 24x36, 24X36, 24 x 36, 24X 36, 24 x36, etc.
        clean = re.sub(r'\s*x\s*', 'x', clean)
    
        # This handles: 24by36, 24BY36, 24 by 36, 24BY 36, 24 by36, 24bY36, 24 bY 36, etc.
        clean = re.sub(r'\s*by\s*', 'x', clean, flags=re.IGNORECASE)
    
        # Now search for pattern: digits + 'x' + digits
        match = re.search(r'(\d+\.?\d*)x(\d+\.?\d*)', clean)
        if not match:
            return ""
        
        w_in = float(match.group(1))
        h_in = float(match.group(2))
    
        # Identify width (smaller) and height (larger)
        if w_in <= h_in:
            width_in, height_in = w_in, h_in
        else:
            width_in, height_in = h_in, w_in

        # --- Width Logic (based on machine) ---
        machine = self.machine_var.get()

        def get_width_ft(inch):
            if machine == "(C.E)":
                if inch <= 26.5: return 2
                elif inch <= 32.5: return 2.5
                elif inch <= 39.5: return 3
                elif inch <= 49.5: return 4
                elif inch <= 60.0: return 5
                else: return 5
            else:  # (C.S)
                if inch <= 26.0: return 2
                elif inch <= 32.0: return 2.5
                elif inch <= 38.0: return 3
                elif inch <= 50.0: return 4
                elif inch <= 62.0: return 5
                elif inch <= 74.0: return 6
                elif inch <= 98.0: return 8
                elif inch <= 124.0: return 10
                else: return 10

        # --- Height Logic (Same for both machines) ---
        def get_height_ft(inch):
            if inch <= 13: return 1
            elif inch <= 18: return 1.5
            elif inch <= 26: return 2
            elif inch <= 32: return 2.5
            elif inch <= 39: return 3
            elif inch <= 44: return 3.5
            elif inch <= 49: return 4
            elif inch <= 53: return 4.5
            elif inch <= 62: return 5
            elif inch <= 68: return 5.5
            elif inch <= 72: return 6
            elif inch <= 78: return 6.5
            elif inch <= 84: return 7
            elif inch <= 90: return 7.5
            elif inch <= 96: return 8
            elif inch <= 103: return 8.5
            elif inch <= 109: return 9
            elif inch <= 114: return 9.5
            elif inch <= 122: return 10
            elif inch <= 126: return 10.5
            elif inch <= 132: return 11
            elif inch <= 138: return 11.5
            elif inch <= 144: return 12
            elif inch <= 151: return 12.5
            elif inch <= 158: return 13
            elif inch <= 163: return 13.5
            elif inch <= 169: return 14
            elif inch <= 175: return 14.5
            elif inch <= 180: return 15
            elif inch <= 186: return 15.5
            elif inch <= 192: return 16
            elif inch <= 198: return 16.5
            elif inch <= 204: return 17
            elif inch <= 210: return 17.5
            elif inch <= 216: return 18
            elif inch <= 222: return 18.5
            elif inch <= 228: return 19
            elif inch <= 234: return 19.5
            elif inch <= 240: return 20
            elif inch <= 246: return 20.5
            elif inch <= 252: return 21
            elif inch <= 258: return 21.5
            elif inch <= 264: return 22
            elif inch <= 270: return 22.5
            elif inch <= 276: return 23
            elif inch <= 282: return 23.5
            elif inch <= 288: return 24
            elif inch <= 294: return 24.6
            elif inch <= 300: return 25
            elif inch <= 306: return 25.5
            elif inch <= 312: return 26
            elif inch <= 318: return 26.5
            elif inch <= 324: return 27
            elif inch <= 330: return 27.5
            elif inch <= 336: return 28
            elif inch <= 342: return 28.5
            elif inch <= 348: return 29
            elif inch <= 354: return 29.5
            elif inch <= 360: return 30
            elif inch <= 366: return 30.5
            elif inch <= 372: return 31
            elif inch <= 378: return 31.5
            elif inch <= 384: return 32
            elif inch <= 390: return 32.5
            elif inch <= 396: return 33
            elif inch <= 402: return 33.5
            elif inch <= 408: return 34
            elif inch <= 414: return 34.5
            elif inch <= 420: return 35
            elif inch <= 426: return 35.5
            elif inch <= 432: return 36
            elif inch <= 438: return 36.5
            elif inch <= 444: return 37
            elif inch <= 450: return 37.5
            elif inch <= 456: return 38
            elif inch <= 462: return 38.5
            elif inch <= 468: return 39
            elif inch <= 474: return 39.5
            elif inch <= 480: return 40
            elif inch <= 486: return 40.5
            elif inch <= 492: return 41
            elif inch <= 498: return 41.5
            elif inch <= 504: return 42
            elif inch <= 510: return 42.5
            elif inch <= 516: return 43
            elif inch <= 522: return 43.5
            elif inch <= 528: return 44
            elif inch <= 534: return 44.5
            elif inch <= 540: return 45
            elif inch <= 546: return 45.5
            elif inch <= 552: return 46
            elif inch <= 558: return 46.5
            elif inch <= 564: return 47
            elif inch <= 570: return 47.5
            elif inch <= 576: return 48
            elif inch <= 582: return 48.5
            elif inch <= 588: return 49
            elif inch <= 594: return 49.5
            elif inch <= 600: return 50
            elif inch <= 606: return 50.5
            elif inch <= 612: return 51
            elif inch <= 618: return 51.5
            elif inch <= 624: return 52
            elif inch <= 630: return 52.5
            elif inch <= 636: return 53
            elif inch <= 642: return 53.5
            elif inch <= 648: return 54
            elif inch <= 654: return 54.5
            elif inch <= 660: return 55
            elif inch <= 666: return 55.5
            elif inch <= 672: return 56
            elif inch <= 678: return 56.5
            elif inch <= 684: return 57
            elif inch <= 690: return 57.5
            elif inch <= 696: return 58
            elif inch <= 702: return 58.5
            elif inch <= 708: return 59
            elif inch <= 714: return 59.5
            elif inch <= 720: return 60
            elif inch <= 726: return 60.5
            elif inch <= 732: return 61
            elif inch <= 738: return 61.5
            elif inch <= 744: return 62
            elif inch <= 750: return 62.5
            elif inch <= 756: return 63
            elif inch <= 762: return 63.5
            elif inch <= 768: return 64
            elif inch <= 774: return 64.5
            elif inch <= 780: return 65
            elif inch <= 786: return 65.5
            elif inch <= 792: return 66
            elif inch <= 798: return 66.5
            elif inch <= 804: return 67
            elif inch <= 810: return 67.5
            elif inch <= 816: return 68
            elif inch <= 822: return 68.5
            elif inch <= 828: return 69
            elif inch <= 834: return 69.5
            elif inch <= 840: return 70
            elif inch <= 846: return 70.5
            elif inch <= 852: return 71
            elif inch <= 858: return 71.5
            elif inch <= 864: return 72
            elif inch <= 870: return 72.5
            elif inch <= 876: return 73
            elif inch <= 882: return 73.5
            elif inch <= 888: return 74
            elif inch <= 894: return 74.5
            elif inch <= 900: return 75
            elif inch <= 906: return 75.5
            elif inch <= 912: return 76
            elif inch <= 918: return 76.5
            elif inch <= 924: return 77
            elif inch <= 930: return 77.5
            elif inch <= 936: return 78
            elif inch <= 942: return 78.5
            elif inch <= 948: return 79
            elif inch <= 954: return 79.5
            elif inch <= 960: return 80
            elif inch <= 966: return 80.5
            elif inch <= 972: return 81
            elif inch <= 978: return 81.5
            elif inch <= 984: return 82
            elif inch <= 990: return 82.5
            elif inch <= 996: return 83
            elif inch <= 1002: return 83.5
            elif inch <= 1008: return 84
            elif inch <= 1014: return 84.5
            elif inch <= 1020: return 85
            elif inch <= 1026: return 85.5
            elif inch <= 1032: return 86
            elif inch <= 1038: return 86.5
            elif inch <= 1044: return 87
            elif inch <= 1050: return 87.5
            elif inch <= 1056: return 88
            elif inch <= 1062: return 88.5
            elif inch <= 1068: return 89
            elif inch <= 1074: return 89.5
            elif inch <= 1080: return 90
            elif inch <= 1086: return 90.5
            elif inch <= 1092: return 91
            elif inch <= 1098: return 91.5
            elif inch <= 1104: return 92
            elif inch <= 1110: return 92.5
            elif inch <= 1116: return 93
            elif inch <= 1122: return 93.5
            elif inch <= 1128: return 94
            elif inch <= 1134: return 94.5
            elif inch <= 1140: return 95
            elif inch <= 1146: return 95.5
            elif inch <= 1152: return 96
            elif inch <= 1158: return 96.5
            elif inch <= 1164: return 97
            elif inch <= 1170: return 97.5
            elif inch <= 1176: return 98
            elif inch <= 1182: return 98.5
            elif inch <= 1188: return 99
            elif inch <= 1194: return 99.5
            elif inch <= 1200: return 100

            else:
                # You'll add more ranges later
                # For now, cap at 100 or extend logic
                return 100  # Placeholder — you'll update this later

        w_ft = get_width_ft(width_in)
        h_ft = get_height_ft(height_in)

        # Enforce minimum area (1x2) — though unlikely here
        if w_ft * h_ft < 2:
            w_ft, h_ft = 1, 2

        # Format: show integers without .0, keep decimals otherwise
        def fmt(val):
            return int(val) if val == int(val) else val

        return f"{fmt(w_ft)}x{fmt(h_ft)}"

    def detect_quantity(self, filename: str) -> int:
        clean = re.sub(r'\s+', ' ', filename.strip().lower())
        clean = re.sub(r'\d+\s*x\s*\d+', '', clean)
        clean = re.sub(r'\s+', ' ', clean).strip()

        for kw in self.quantity_keywords:
            match = re.search(rf'\b(\d+)\s*{re.escape(kw)}\b', clean)
            if match:
                return int(match.group(1))
            match = re.search(rf'\b{re.escape(kw)}\s*(\d+)\b', clean)
            if match:
                return int(match.group(1))
        return 1

    def generate_new_filename(self, original_stem: str, party_code: str, extension: str, dim_str: str = "") -> str:
        qty = self.detect_quantity(original_stem)
        machine = self.machine_var.get()
        dim_part = f"(FT.{dim_str})" if dim_str else ""
        return f"{party_code}_{original_stem} {machine}{dim_part}(Q.{qty})%%{extension}"

    def update_preview(self):
        if not self.selected_file or not self.selected_file.exists():
            self.preview_label.configure(text="Preview: --", text_color="gray")
            return
        stem = self.selected_file.stem
        ext = self.selected_file.suffix
        # Find party folder
        party_folder = self.find_party_folder(self.selected_file)
        if not party_folder:
            color = "red"
            code = "?"
        else:
            code = self.get_party_code(party_folder.name)
            color = "red" if code == "?" else "lightgreen"
        dim = self.extract_dimensions(self.selected_file.name)
        new_name = self.generate_new_filename(stem, code, ext, dim)
        self.preview_label.configure(text=f"Preview: {new_name}", text_color=color)

    def select_folder(self):
        try:
            folder = filedialog.askdirectory(title="Select Year Folder (e.g. 2025)")
            if not folder: return
            path = Path(folder)
            if not path.is_dir(): return
            self.selected_root = path
            self.update_folder_path_display()
            self.last_folder = str(folder)
            self.save_config()
            self.scan_folder()
            self.start_auto_scan()
            self.status_label.configure(text=f"📁 Active: {path.name}")
        except Exception as e:
            self.status_label.configure(text="❌ Folder selection failed")
            messagebox.showerror("Error", f"Could not open folder:\n{e}")

    def scan_folder(self):
        # Show loading message
        self.status_label.configure(text="🔍 Scanning folder... Please wait")
        # Disable buttons
        self.scan_btn.configure(state="disabled")
        self.reload_csv_btn.configure(state="disabled")
        self.select_all_btn.configure(state="disabled")
        self.rename_btn.configure(state="disabled")
        # Update UI immediately
        self.update_idletasks()
        if not self.selected_root or not self.selected_root.exists():
            self.file_listbox.delete("1.0", "end")  # Use correct index for tk.Text
            header_text = f"{'Code':<6} {'Party':<15} Filename\n{'-'*60}\n"
            self.file_listbox.insert("end", header_text)
            return
        try:
            files = []
            for ext in self.allowed_extensions:
                for file_path in self.selected_root.rglob(f"*{ext}"):
                    # Find the closest party folder
                    party_folder = self.find_party_folder(file_path)
                    if not party_folder:
                        continue  # Not under any party folder

                    # Check if in Done folder
                    in_done = "done" in [p.lower() for p in file_path.parts]

                    # Finalize Mode: Only show files in Done and NOT finalized
                    if self.show_done_var.get():
                        if not in_done:
                            continue  # Not in Done → skip
                        if "[ok]" in file_path.name:
                            continue  # Already finalized → skip
                    else:
                        # Normal Mode: Skip files in Done
                        if in_done:
                            continue
                        # Skip already finalized files
                        if "[ok]" in file_path.name:
                            continue

                    files.append(file_path)

            # Sort and update UI
            files = sorted(files, key=lambda x: x.name.lower())
            self.file_listbox.delete("0.0", "end")

            mode = "Finalize Mode" if self.show_done_var.get() else "New Files"
            self.file_listbox.insert("0.0", f"📁 {mode}\nFiles:\n")

            self.file_path_list = []
            for file_path in files:
                party_folder = self.find_party_folder(file_path)
                if not party_folder:
                    continue

                # Find the original party name as it appears in parties.csv
                original_party_name = next(
                    (name for name in self.party_map if name.lower() == party_folder.name.lower()),
                    party_folder.name
                )
                code = self.party_map.get(original_party_name, "?")
                display = f"{code} | {original_party_name} | {file_path.name}\n"
                self.file_listbox.insert("end", display, "normal")
                self.file_path_list.append(file_path)

            self.filtered_file_list = self.file_path_list[:]

            # Update UI
            if self.show_done_var.get():
                self.select_all_btn.grid_remove()
            else:
                self.select_all_btn.grid()

            self.status_label.configure(text=f"✅ {mode}: {len(self.file_path_list)} files")
            self.update_info_bar()
            self.update_folder_path_display()
            # Re-enable buttons
            self.scan_btn.configure(state="normal")
            self.reload_csv_btn.configure(state="normal")
            self.select_all_btn.configure(state="normal")
            self.rename_btn.configure(state="normal")

        except Exception as e:
            self.status_label.configure(text=f"❌ Scan error: {e}")

    def on_search_change(self, *args):
        query = self.search_var.get().strip().lower()
        self.file_listbox.delete("1.0", "end")

        # Clear and rebuild filtered list
        if query == "":
            self.filtered_file_list = self.file_path_list.copy()
        else:
            self.filtered_file_list.clear()
            for file_path in self.file_path_list:
                if query in file_path.name.lower():
                    self.filtered_file_list.append(file_path)

        # Sort for consistency
        self.filtered_file_list.sort(key=lambda x: x.name.lower())

        # Build header with proper formatting and newline
        header_text = f"{'Code':<6} {'Party':<15} Filename\n{'-'*60}\n"
        self.file_listbox.insert("end", header_text)

        # Insert files
        for file_path in self.filtered_file_list:
            party_folder = self.find_party_folder(file_path)
            if not party_folder:
                continue
            original_party_name = next(
                (name for name in self.party_map if name.lower() == party_folder.name.lower()),
                party_folder.name
            )
            code = self.party_map.get(original_party_name, "?")
            display = f"{code:<6} {original_party_name:<15} {file_path.name}\n"
            self.file_listbox.insert("end", display)

        # Update status
        self.status_label.configure(text=f"🔍 Found {len(self.filtered_file_list)} matching files")

    def on_file_click(self, event):
        try:
            index = self.file_listbox.index(f"@{event.x},{event.y}")
            line_num = int(index.split('.')[0])
            HEADER_LINES = 3
            file_index = line_num - HEADER_LINES
            if file_index < 0 or file_index >= len(self.filtered_file_list): return
            selected_path = self.filtered_file_list[file_index]
            if not selected_path.exists(): return
            self.selected_file = selected_path.resolve()

            # Only highlight in Normal Mode
            if not self.show_done_var.get():
                # Remove previous highlights
                self.file_listbox.tag_remove("selected", "1.0", "end")
                # Apply highlight to current line
                start_line = f"{line_num}.0"
                end_line = f"{line_num}.end"
                self.file_listbox.tag_add("selected", start_line, end_line)
                self.file_listbox.see(start_line)
            else:
                # Finalize Mode: Open popup
                self.open_manual_finalize_popup()
                return

            # Update UI
            self.file_label.configure(text=f"Selected: {selected_path.name}")
            self.update_preview()
            self.status_label.configure(text=f"📁 {selected_path.parent.name} | Ready")

            # Flash effect
            self.file_label.configure(text_color="lightblue")
            self.after(100, lambda: self.file_label.configure(text_color="white"))

        except Exception as e:
            self.status_label.configure(text=f"❌ Click error: {e}")

    def open_manual_finalize_popup(self):
        if not self.selected_file or "[ok]" in self.selected_file.name:
            return

        files_left = [p for p in self.filtered_file_list if "[ok]" not in str(p)]
        try:
            current_idx = files_left.index(self.selected_file)
        except (ValueError, AttributeError):
            current_idx = 0

        popup = ctk.CTkToplevel(self)
        popup.title("✅ Finalize File")
        popup.geometry("700x350")
        popup.resizable(False, False)
        popup.transient(self)
        popup.grab_set()

        # Center popup
        x = self.winfo_x() + (self.winfo_width() // 2) - 200
        y = self.winfo_y() + (self.winfo_height() // 2) - 115
        popup.geometry(f"+{int(x)}+{int(y)}")

        # Top-right "Open Folder" button
        folder_btn = ctk.CTkButton(
            popup,
            text="📁 Open Folder",
            command=lambda: os.startfile(str(self.selected_file.parent)),
            width=100,
            height=24,
            font=("Segoe UI", 12)
        )
        folder_btn.place(relx=1.0, rely=0.0, x=-10, y=10, anchor="ne")

        # Show file name
        ctk.CTkLabel(popup, text="📄 File:", font=("Segoe UI", 11, "bold")).pack(pady=(10, 2), anchor="w", padx=20)
        ctk.CTkLabel(
            popup,
            text=self.selected_file.name,
            font=("Consolas", 16),
            text_color="lightblue"
        ).pack(pady=(0, 10), padx=20, anchor="w")

        # Show party name and code
        party_folder = self.find_party_folder(self.selected_file)
        if party_folder:
            party_name = party_folder.name
            party_code = self.get_party_code(party_name)
            ctk.CTkLabel(
                popup,
                text=f"📁 Party: {party_name} (Code: {party_code})",
                font=("Segoe UI", 15, "bold"),
                text_color="orange"
            ).pack(pady=(0, 10), padx=20, anchor="w")
        else:
            ctk.CTkLabel(
                popup,
                text="📁 Party: Not found",
                font=("Segoe UI", 15, "bold"),
                text_color="Red"
            ).pack(pady=(0, 10), padx=20, anchor="w")

        # Extract current values
        stem = self.selected_file.stem
        ext = self.selected_file.suffix

        # Detect FT
        ft_match = re.search(r'\(FT\.(\d+\.?\d*x\d+\.?\d*)\)', stem)
        ft_value = ft_match.group(1) if ft_match else "1x1"

        # Detect Q
        qty_match = re.search(r'\(Q\.(\d+)\)', stem)
        qty_value = qty_match.group(1) if qty_match else "1"

        # Detect Cat
        cat_match = re.search(r'%(\d+)%', stem)
        cat_value = cat_match.group(1) if cat_match else "1"

        # Form layout
        frame = ctk.CTkFrame(popup, fg_color="transparent")
        frame.pack(pady=6)

        # FT Entry
        ctk.CTkLabel(frame, text="FT:").pack(side="left", padx=5)
        ft_var = ctk.StringVar(value=ft_value)
        ft_entry = ctk.CTkEntry(frame, textvariable=ft_var, width=100)
        ft_entry.pack(side="left", padx=10)

        # Cat Entry
        ctk.CTkLabel(frame, text="Cat:").pack(side="left", padx=5)
        cat_var = ctk.StringVar(value="1")
        cat_entry = ctk.CTkEntry(frame, textvariable=cat_var, width=60)
        cat_entry.pack(side="left", padx=10)

        # Q Entry
        ctk.CTkLabel(frame, text="Qty:").pack(side="left", padx=5)
        qty_var = ctk.StringVar(value=qty_value)
        qty_entry = ctk.CTkEntry(frame, textvariable=qty_var, width=60)
        qty_entry.pack(side="left", padx=10)

        def finalize():
            ft = ft_var.get().strip()
            qty = qty_var.get().strip()
            cat = cat_var.get().strip()

            if not ft or not qty.isdigit() or int(qty) < 1 or not cat:
                messagebox.showwarning("Invalid", "All fields are required and Qty ≥ 1")
                return

            try:
                old_path = self.selected_file
                if "[ok]" in old_path.name:
                    return
            except Exception as e:
                messagebox.showerror("Error", f"Failed to access file: {e}")
                return

            base = old_path.stem
            base = re.sub(r'\(FT\.\d+\.?\d*x\d+\.?\d*\)', '', base)
            base = re.sub(r'\(Q\.\d+\)', '', base)
            base = re.sub(r'%\d+%', '', base)
            base = re.sub(r'%%+', '', base)
            base = re.sub(r'\[ok\]', '', base)
            base = re.sub(r'\s+', ' ', base).strip()

            machine = self.machine_var.get()
            if machine in base and "[ok]" not in base:
                base = base.replace(machine, f"{machine}")
                

                new_name = f"{base}(FT.{ft})(Q.{qty})%{cat}%[ok]{ext}"
                new_path = old_path.parent / new_name

                # Handle duplicate filenames
                counter = 1
                original_new_path = new_path
                while new_path.exists():
                    name_without_ext = f"{original_new_path.stem} ({counter})"
                    new_path = original_new_path.parent / f"{name_without_ext}{ext}"
                    counter += 1
                    if counter > 100:
                        messagebox.showerror("Error", "Too many duplicates. Clean up folder.")
                        return

                # Rename
                old_path.rename(new_path)
                self.history.add(old_path, new_path)
                self.status_label.configure(text=f"✅ Finalized: {new_path.name}")

                # Update last finalized preview
                if hasattr(popup, 'last_finalized_label'):
                    popup.last_finalized_label.configure(text=f"Last: {new_path.name}")
                    popup.last_finalized_path_label.configure(text=f"Path: {new_path.parent}")

                # Auto-advance to next file
                popup.destroy()
                self.scan_folder()

                # Recalculate list after rename
                files_left_after = [p for p in self.filtered_file_list if "[ok]" not in str(p)]

                # Find the current file's new index in the updated list
                try:
                    current_idx = files_left_after.index(self.selected_file)
                except ValueError:
                    # Current file not found (already finalized or moved)
                    if files_left_after:
                        self.selected_file = files_left_after[0]
                        self.after(100, self.open_manual_finalize_popup)
                    else:
                        self.status_label.configure(text="✅ All files finalized")
                    return

                # Move to next file
                next_idx = current_idx + 1
                if next_idx < len(files_left_after):
                    self.selected_file = files_left_after[next_idx]
                    self.after(100, self.open_manual_finalize_popup)
                else:
                    self.status_label.configure(text="✅ All files finalized")

        # Buttons
        btn_frame = ctk.CTkFrame(popup, fg_color="transparent")
        btn_frame.pack(pady=12)

        # OK Button (focusable)
        ok_btn = ctk.CTkButton(btn_frame, text="✅ OK", command=finalize, width=80)
        ok_btn.pack(side="left", padx=5)

        # Cancel Button
        ctk.CTkButton(btn_frame, text="Cancel", command=popup.destroy, width=80).pack(side="left", padx=5)

        # Status bar for last finalized preview (bottom)
        status_frame = ctk.CTkFrame(popup, fg_color="transparent")
        status_frame.pack(side="bottom", fill="x", padx=20, pady=10)

        # Label for filename
        popup.last_finalized_label = ctk.CTkLabel(
            status_frame,
            text="Last: —",
            font=("Consolas", 12),
            text_color="gray",
            anchor="w"
        )
        popup.last_finalized_label.pack(fill="x")

        # Label for path
        popup.last_finalized_path_label = ctk.CTkLabel(
            status_frame,
            text="",
            font=("Consolas", 10),
            text_color="gray",
            anchor="w"
        )
        popup.last_finalized_path_label.pack(fill="x", pady=(2, 0))

        # Tab Navigation with Text Selection
        def select_on_focus(widget, var):
            def on_focus_in(e):
                widget.select_range(0, "end")
                widget.icursor("end")
            widget.bind("<FocusIn>", on_focus_in)

        select_on_focus(ft_entry, ft_var)
        select_on_focus(cat_entry, cat_var)
        select_on_focus(qty_entry, qty_var)

        # Tab Order
        ft_entry.bind("<Tab>", lambda e: (cat_entry.focus(), "break"))
        cat_entry.bind("<Tab>", lambda e: (qty_entry.focus(), "break"))
        qty_entry.bind("<Tab>", lambda e: (ok_btn.focus(), "break"))

        # Enter Key Support
        qty_entry.bind("<Return>", lambda e: finalize())
        ok_btn.bind("<Return>", lambda e: finalize())
        ok_btn.bind("<KP_Enter>", lambda e: finalize())

        # Focus first field
        ft_entry.focus()

    def get_party_code(self, folder_name: str) -> str:
        """Get party code case-insensitively"""
        if not folder_name:
            return "?"
        for name in self.party_map:
            if name.lower() == folder_name.lower():
                return self.party_map[name]
        return "?"

    def show_manual_input_popup(self, callback, filename=None):
        popup = ctk.CTkToplevel(self)
        popup.title("✏️ Finalize File")
        popup.geometry("700x350")
        popup.resizable(False, False)
        popup.transient(self)
        popup.grab_set()

        x = self.winfo_x() + (self.winfo_width() // 2) - 190
        y = self.winfo_y() + (self.winfo_height() // 2) - 105
        popup.geometry(f"+{int(x)}+{int(y)}")

        if filename:
            ctk.CTkLabel(popup, text="📄 File:", font=("Helvetica", 11, "bold")).pack(pady=(10, 2), anchor="w", padx=20)
            ctk.CTkLabel(popup, text=filename, font=("Consolas", 16), text_color="lightblue").pack(pady=(0, 8), padx=20)

        frame = ctk.CTkFrame(popup, fg_color="transparent")
        frame.pack(pady=4)

        ctk.CTkLabel(frame, text="Qty:").pack(side="left", padx=5)
        qty_var = ctk.StringVar(value="1")
        qty_entry = ctk.CTkEntry(frame, textvariable=qty_var, width=60)
        qty_entry.pack(side="left", padx=10)
        qty_entry.focus()

        ctk.CTkLabel(frame, text="Cat:").pack(side="left", padx=5)
        cat_var = ctk.StringVar()
        cat_entry = ctk.CTkEntry(frame, textvariable=cat_var, width=60)
        cat_entry.pack(side="left", padx=10)

        btn_frame = ctk.CTkFrame(popup, fg_color="transparent")
        btn_frame.pack(pady=12)

        def submit():
            qty = qty_var.get().strip()
            if not qty.isdigit() or int(qty) < 1:
                messagebox.showwarning("Invalid", "Quantity must be ≥ 1")
                return
            callback(qty, cat_var.get().strip())
            popup.destroy()
            self.scan_folder()  # Refresh list

            # Recalculate list after rename
            files_left_after = [p for p in self.filtered_file_list if "[ok]" not in str(p)]

            # Use stored index to find next file
            next_idx = popup.current_idx + 1
            if next_idx < len(files_left_after):
                self.selected_file = files_left_after[next_idx]
                self.after(100, self.open_manual_finalize_popup)
            else:
                self.status_label.configure(text="✅ All files finalized")
            

        def cancel():
            popup.destroy()

        ctk.CTkButton(btn_frame, text="✅ OK", command=submit, width=80).pack(side="left", padx=5)
        ctk.CTkButton(btn_frame, text="Cancel", command=cancel, width=80).pack(side="left", padx=5)

        popup.bind("<Return>", lambda e: submit())
        popup.bind("<Escape>", lambda e: cancel())
        qty_entry.bind("<Tab>", lambda e: cat_entry.focus())
        cat_entry.bind("<Return>", lambda e: submit())

    def rename_file(self):
        if not self.selected_file: return
        machine = self.machine_var.get()
        if not machine: return
        file_path = self.selected_file
        if not file_path.exists(): return
        # Find the party folder
        party_folder = self.find_party_folder(file_path)
        if not party_folder:
            self.status_label.configure(text="❌ Not in a party folder")
            return
        code = self.get_party_code(party_folder.name)
        if code == "?":
            self.status_label.configure(text="❌ Party code not found")
            return
        stem = file_path.stem
        ext = file_path.suffix
        dim = self.extract_dimensions(file_path.name)
        new_name = self.generate_new_filename(stem, code, ext, dim)
        new_path = file_path.parent / new_name
        counter = 1
        original = new_path
        while new_path.exists():
            new_path = original.parent / f"{original.stem} ({counter}){ext}"
            counter += 1
        try:
            file_path.rename(new_path)
            done_folder = file_path.parent / "Done"
            done_folder.mkdir(exist_ok=True)
            final_path = done_folder / new_path.name
            shutil.move(str(new_path), str(final_path))
            self.history.add(file_path, final_path)
            self.status_label.configure(text=f"✅ Renamed: {final_path.name}")
            self.update_info_bar()
            self.scan_folder()
            self.selected_file = None
            self.file_label.configure(text="No file selected")
            self.preview_label.configure(text="Preview: --", text_color="gray")
        except Exception as e:
            self.status_label.configure(text=f"❌ Error: {e}")

    def select_all_files(self):
        # Show processing
        self.status_label.configure(text="📦 Batch renaming... Please wait")
        # Disable buttons
        self.select_all_btn.configure(state="disabled")
        self.scan_btn.configure(state="disabled")
        self.rename_btn.configure(state="disabled")
        self.undo_btn.configure(state="disabled")
        self.redo_btn.configure(state="disabled")
        self.update_idletasks()

        # Check if we should proceed
        if not self.file_path_list or self.show_done_var.get():
            self.status_label.configure(text="❌ No files to rename in normal mode")
            self.select_all_btn.configure(state="normal")
            self.scan_btn.configure(state="normal")
            self.rename_btn.configure(state="normal")
            self.undo_btn.configure(state="normal")
            self.redo_btn.configure(state="normal")
            return

        machine = self.machine_var.get()
        if not machine:
            self.status_label.configure(text="❌ Machine not selected")
            self.select_all_btn.configure(state="normal")
            self.scan_btn.configure(state="normal")
            self.rename_btn.configure(state="normal")
            self.undo_btn.configure(state="normal")
            self.redo_btn.configure(state="normal")
            return

        confirm = messagebox.askyesno("Confirm Batch Rename", f"Rename all {len(self.file_path_list)} files?\nThis cannot be undone.")
        if not confirm:
            self.select_all_btn.configure(state="normal")
            self.scan_btn.configure(state="normal")
            self.rename_btn.configure(state="normal")
            self.undo_btn.configure(state="normal")
            self.redo_btn.configure(state="normal")
            return

        # ====== Progress Popup ======
        progress_popup = ctk.CTkToplevel(self)
        progress_popup.title("Batch Rename Progress")
        progress_popup.geometry("500x180")
        progress_popup.resizable(False, False)
        progress_popup.transient(self)
        progress_popup.grab_set()
        progress_popup.protocol("WM_DELETE_WINDOW", lambda: None)  # Prevent closing

        # Center popup
        x = self.winfo_x() + (self.winfo_width() // 2) - 250
        y = self.winfo_y() + (self.winfo_height() // 2) - 90
        progress_popup.geometry(f"+{int(x)}+{int(y)}")

        ctk.CTkLabel(progress_popup, text="Batch Renaming Files...", font=TITLE_FONT).pack(pady=(10, 5))
            # ====== Random Fun Messages ======
        funny_messages = [
           "🚀 ફાઇલ છૂટી! જેલમાંથી બહાર!"
        ]

        # ====== Status Label and Progress Bar ======
        status_label = ctk.CTkLabel(progress_popup, text="Starting...", font=FONT, wraplength=450)
        status_label.pack(pady=5)

        progress_bar = ctk.CTkProgressBar(progress_popup, width=400)
        progress_bar.pack(pady=10)

        # ====== Random Message Label (Red Area) ======
        random_msg_label = ctk.CTkLabel(
            progress_popup,
            text="",
            font=("Segoe UI", 18),
            text_color="white",
            bg_color="red",
            fg_color="red",
            corner_radius=6,
            height=35,
            anchor="w"
        )
        random_msg_label.pack(pady=10)

        # ====== Update Random Message After Each File ======
        def update_random_message():
            if not self.file_path_list:
                return
            msg = self.random.choice(funny_messages)
            random_msg_label.configure(text=msg)

        # Import random inside function
        import random
        self.random = random  # Store for use in inner function

        # ====== Variables for Iterative Processing ======
        file_list = self.file_path_list[:]
        total = len(file_list)
        renamed_count = 0
        error_count = 0

        # ====== Define Step-by-Step Function ======
        def process_next_file():
            nonlocal renamed_count, error_count

            if not file_list:
                # === FINISH ===
                progress_popup.destroy()
                self.scan_folder()
                self.status_label.configure(text=f"✅ Batch: {renamed_count} renamed, {error_count} failed")
                # 🎵 Play sound only if file exists
                if SOUND_FILE.exists():
                    try:
                        winsound.PlaySound(str(SOUND_FILE), winsound.SND_FILENAME | winsound.SND_ASYNC)
                    except Exception as e:
                        print("Sound error:", e)

                # 🎉 Trigger Easter Egg: Jump to fun messages
                self.current_msg_index = len(self.bulletin_messages) - 5  # Start from first easter egg
                self.bulletin_offset = 60 + 50  # visible_width + message length (start off-screen right)
                # Re-enable buttons
                self.select_all_btn.configure(state="normal")
                self.scan_btn.configure(state="normal")
                self.rename_btn.configure(state="normal")
                self.undo_btn.configure(state="normal")
                self.redo_btn.configure(state="normal")
                self.update_info_bar()
                return

            file_path = file_list.pop(0)

            try:
                if not file_path.exists():
                    status_label.configure(text=f"⚠️ Not found: {file_path.name}")
                    error_count += 1
                else:
                    party_folder = self.find_party_folder(file_path)
                    if not party_folder:
                        status_label.configure(text=f"⚠️ No party: {file_path.name}")
                        error_count += 1
                    else:
                        code = self.get_party_code(party_folder.name)
                        if code == "?":
                            status_label.configure(text=f"⚠️ No code: {file_path.name}")
                            error_count += 1
                        else:
                            stem = file_path.stem
                            ext = file_path.suffix
                            dim = self.extract_dimensions(file_path.name)
                            new_name = self.generate_new_filename(stem, code, ext, dim)
                            new_path = file_path.parent / new_name

                            # Handle conflict
                            counter = 1
                            while new_path.exists():
                                new_path = file_path.parent / f"{new_name[:-len(ext)]} ({counter}){ext}"
                                counter += 1

                            # Rename + Move to Done
                            file_path.rename(new_path)
                            done_folder = file_path.parent / "Done"
                            done_folder.mkdir(exist_ok=True)
                            final_path = done_folder / new_path.name
                            shutil.move(str(new_path), str(final_path))
                            self.history.add(file_path, final_path)
                            renamed_count += 1
                            status_label.configure(text=f"✅ Renamed: {file_path.name}")
                            # Update random message in red area
                            update_random_message()  # ← Add this line here

            except Exception as e:
                error_count += 1
                status_label.configure(text=f"❌ Error: {file_path.name}")
                logging.error(f"Batch rename failed for {file_path}: {e}")

            # Update progress
            progress = renamed_count + error_count
            progress_bar.set(progress / total)
            progress_popup.update_idletasks()

            # Schedule next file
            self.after(50, process_next_file)

        # ====== Start Processing ======
        self.after(100, process_next_file)

    def undo_all_batch(self):
        if not self.history.history:
            messagebox.showinfo("Undo All", "Nothing to undo.")
            return

        confirm = messagebox.askyesno("Undo All", f"Undo {len(self.history.history)} renamed files?\nThis cannot be undone.")
        if not confirm:
            return

        # ====== Progress Popup ======
        progress_popup = ctk.CTkToplevel(self)
        progress_popup.title("Undo All Progress")
        progress_popup.geometry("500x180")
        progress_popup.resizable(False, False)
        progress_popup.transient(self)
        progress_popup.grab_set()
        progress_popup.protocol("WM_DELETE_WINDOW", lambda: None)  # Prevent closing

        # Center popup
        x = self.winfo_x() + (self.winfo_width() // 2) - 250
        y = self.winfo_y() + (self.winfo_height() // 2) - 90
        progress_popup.geometry(f"+{int(x)}+{int(y)}")

        ctk.CTkLabel(progress_popup, text="Undoing All Renames...", font=TITLE_FONT).pack(pady=(10, 5))
        status_label = ctk.CTkLabel(progress_popup, text="Starting...", font=FONT, wraplength=450)
        status_label.pack(pady=5)

        progress_bar = ctk.CTkProgressBar(progress_popup, width=400)
        progress_bar.set(0)
        progress_bar.pack(pady=10)

        # ====== COPY OF BATCH RENAME'S RED AREA ======
        # ====== Random Fun Messages (Same as Batch Rename) ======
        funny_messages = [
            "🚀 તો ડોફા પેલા જોય ને કરતો હોય તો...!"
        ]

        import random
        self.random = random

        def update_random_message():
            msg = self.random.choice(funny_messages)
            random_msg_label.configure(text=msg)

        random_msg_label = ctk.CTkLabel(
            progress_popup,
            text="",
            font=("Consolas", 18, "bold"),
            text_color="white",
            bg_color="red",
            fg_color="red",
            corner_radius=6,
            height=28,
            anchor="w"
        )
        random_msg_label.pack(pady=10)

        # ====== Prepare ======
        items = self.history.history[:]
        total = len(items)
        restored = 0

        # ====== Step-by-Step Undo ======
        def undo_next():
            nonlocal restored

            if not items:
                # === FINISH ===
                progress_popup.destroy()
                self.scan_folder()
                self.status_label.configure(text=f"↩ Undo All: {restored} files restored")
                self.update_info_bar()
                return

            item = items.pop(0)
            try:
                src = Path(item["new"])
                dst = Path(item["old"])
                dst.parent.mkdir(exist_ok=True)

                # Move file back
                shutil.move(str(src), str(dst))
                # Remove from history
                self.history.undo()
                restored += 1
                status_label.configure(text=f"↩ Restored: {dst.name}")

                # Update red area message
                update_random_message()

            except Exception as e:
                status_label.configure(text=f"❌ Failed: {dst.name}")
                logging.error(f"Undo failed: {e}")
                update_random_message()

            # Update progress
            progress_bar.set(restored / total)
            progress_popup.update_idletasks()

            # Schedule next
            self.after(50, undo_next)

        # ====== Start Processing ======
        self.after(100, undo_next)

    def undo_rename(self):
            item = self.history.undo()
            if not item: return
            try:
                src = Path(item["new"])
                dst = Path(item["old"])
                dst.parent.mkdir(exist_ok=True)
                shutil.move(str(src), str(dst))
                self.status_label.configure(text=f"↩ Undo: {dst.name}")
                self.update_info_bar()
                self.scan_folder()
            except Exception as e:
                self.status_label.configure(text=f"❌ Undo failed: {e}")

    def redo_rename(self):
        item = self.history.redo()
        if not item: return
        try:
            src = Path(item["old"])
            dst = Path(item["new"])
            counter = 1
            orig = dst
            while dst.exists():
                dst = orig.parent / f"{orig.stem} ({counter}){orig.suffix}"
                counter += 1
            src.rename(dst)
            done = dst.parent / "Done"
            done.mkdir(exist_ok=True)
            shutil.move(str(dst), str(done / dst.name))
            self.status_label.configure(text=f"⟳ Redo: {dst.name}")
            self.update_info_bar()
            self.scan_folder()
        except Exception as e:
            self.status_label.configure(text=f"❌ Redo failed: {e}")

    def load_parties_csv(self):
        # Show loading
        self.status_label.configure(text="🔄 Loading parties CSV...")
        # Disable buttons
        self.reload_csv_btn.configure(state="disabled")
        self.scan_btn.configure(state="disabled")
        self.select_all_btn.configure(state="disabled")
        # Update UI
        self.update_idletasks()
        csv_path = codes_dir / "parties.csv"
        if not csv_path.exists():
            try:
                # Instead of writing manually, use the save function
                self.party_map = {
                    "Creative": "2",
                    "Pranam Maheta": "7",
                    "XYZ Designs": "5",
                    "Sunrise": "3",
                    "Vikas": "9"
                }
                self.save_parties_csv()
                self.status_label.configure(text=f"✅ Created default CSV: {csv_path.name}")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to create CSV: {e}")
                self.party_map = {}
                return

        # Load CSV
        self.party_map = {}
        try:
            with open(csv_path, mode='r', encoding='utf-8') as f:
                reader = csv.DictReader(f)
                for row in reader:
                    name = row.get("Party Name", "").strip()
                    code = row.get("Code", "").strip()
                    if name and code:
                        self.party_map[name] = code
            self.status_label.configure(text=f"✅ Loaded {len(self.party_map)} parties from {csv_path.name}")
            self.update_info_bar()
            self.reload_csv_btn.configure(state="normal")
            self.scan_btn.configure(state="normal")
            self.select_all_btn.configure(state="normal")
            if self.selected_root:
                self.scan_folder()
        except Exception as e:
            self.status_label.configure(text="❌ Failed to load CSV")
            messagebox.showerror("Error", f"Failed to load parties.csv:\n{e}")

    def load_config(self):
        try:
            with open(self.config_file, "r", encoding="utf-8") as f:
                data = json.load(f)
                self.last_folder = data.get("last_folder", "")
        except Exception:
            self.last_folder = ""

    def find_party_folder(self, path):
        """Find the closest parent folder that is a known party (case-insensitive)"""
        current = path.parent
        while True:
            # Check if this folder is a known party
            if current.name.lower() in [p.lower() for p in self.party_map]:
                return current
            # Stop if we've reached the root
            if str(current) == str(current.parent):
                break
            current = current.parent
        return None

    def save_config(self):
        try:
            with open(self.config_file, "w", encoding="utf-8") as f:
                json.dump({"last_folder": self.last_folder}, f, indent=2)
        except Exception as e:
            logging.error(f"Config save error: {e}")

    def toggle_auto_scan(self):
        if self.auto_scan_var.get() and self.selected_root:
            self.start_auto_scan()
        else:
            self.stop_auto_scan()

    def start_auto_scan(self):
        self.stop_auto_scan()
        if not self.selected_root:
            return

        class Handler(FileSystemEventHandler):
            def __init__(self, app):
                self.app = app
            def on_created(self, event):
                if not event.is_directory:
                    ext = Path(event.src_path).suffix.lower()
                    if ext in self.app.allowed_extensions:
                        self.app.after(0, self.app.scan_folder)

        handler = Handler(self)  # ✅ Define handler here
        self.auto_observer = Observer()
        self.auto_observer.schedule(handler, str(self.selected_root), recursive=True)
        self.auto_observer.start()

    def stop_auto_scan(self):
        if self.auto_observer:
            self.auto_observer.stop()
            self.auto_observer.join()
            self.auto_observer = None

    def destroy(self):
        self.stop_auto_scan()
        self.save_config()
        super().destroy()

    def open_parties_editor(self):
        """Open UI to add/edit/delete parties without flicker or errors"""
        # === Create popup and hide it until ready ===
        popup = ctk.CTkToplevel(self)
        popup.title("👥 Manage Party Codes")
        popup.geometry("500x500")
        popup.resizable(False, False)
        popup.transient(self)
        popup.grab_set()
        popup.withdraw()  # Hide during setup

        # Center popup
        self.update_idletasks()
        x = self.winfo_x() + (self.winfo_width() // 2) - 250
        y = self.winfo_y() + (self.winfo_height() // 2) - 250
        popup.geometry(f"+{int(x)}+{int(y)}")

        # === Header ===
        ctk.CTkLabel(popup, text="Manage Party Codes", font=("Segoe UI", 18, "bold")).pack(pady=10)

        # === Search Frame ===
        search_frame = ctk.CTkFrame(popup, fg_color="transparent")
        search_frame.pack(pady=(5, 0), padx=20, fill="x")

        ctk.CTkLabel(search_frame, text="🔍 Search:", font=("Segoe UI", 11)).pack(side="left")
        search_var = ctk.StringVar()
        search_var.trace_add("write", lambda *args: refresh_list())  # ✅ Re-enable live search
        search_entry = ctk.CTkEntry(search_frame, textvariable=search_var, placeholder_text="Filter parties...")
        search_entry.pack(side="left", fill="x", expand=True, padx=(5, 0))

        # === Scrollable List ===
        list_frame = ctk.CTkScrollableFrame(popup, height=250)
        list_frame.pack(pady=10, padx=20, fill="both", expand=True)
        party_widgets = []  # Store (name, entry, delete_btn) to access later

        def refresh_list():
            # Clear existing widgets
            for widget in list_frame.winfo_children():
                widget.destroy()
            party_widgets.clear()

            query = search_var.get().strip().lower()
            filtered_parties = sorted([name for name in self.party_map.keys() if not query or query in name.lower()])

            # Group by first letter
            groups = {}
            for name in filtered_parties:
                first_letter = name[0].upper()
                if first_letter not in groups:
                    groups[first_letter] = []
                groups[first_letter].append(name)

            # Add groups to UI
            for letter in sorted(groups.keys()):
                ctk.CTkLabel(
                    list_frame,
                    text=letter,
                    font=("Segoe UI", 12, "bold"),
                    text_color="cyan"
                ).pack(anchor="w", padx=10, pady=(5, 0))

                for name in groups[letter]:
                    inner_frame = ctk.CTkFrame(list_frame)
                    inner_frame.pack(fill="x", pady=2)

                    ctk.CTkLabel(inner_frame, text=name, width=180, anchor="w").pack(side="left", padx=5)

                    code_entry = ctk.CTkEntry(inner_frame, width=60, placeholder_text="Code")
                    code_entry.insert(0, self.party_map[name])
                    code_entry.pack(side="left")

                    def make_delete_handler(party_name=name, frame=inner_frame):
                        def handler():
                            if messagebox.askyesno("Delete Party", f"Delete '{party_name}'?"):
                                del self.party_map[party_name]
                                frame.destroy()
                        return handler

                    del_btn = ctk.CTkButton(
                        inner_frame,
                        text="🗑️",
                        width=40,
                        height=28,
                        fg_color="#d42222",
                        hover_color="#a00",
                        command=make_delete_handler(name, inner_frame)
                    )
                    del_btn.pack(side="right")
                    party_widgets.append((name, code_entry, del_btn))

        refresh_list()

        # === Add New Party ===
        add_frame = ctk.CTkFrame(popup)
        add_frame.pack(pady=10, padx=20, fill="x")

        ctk.CTkLabel(add_frame, text="Add New:").pack(side="left")
        add_name_entry = ctk.CTkEntry(add_frame, placeholder_text="Party Name")
        add_name_entry.pack(side="left", fill="x", expand=True, padx=5)
        add_code_entry = ctk.CTkEntry(add_frame, placeholder_text="Code", width=80)
        add_code_entry.pack(side="left")

        def add_party():
            name = add_name_entry.get().strip()
            code = add_code_entry.get().strip()
            if not name:
                messagebox.showwarning("Empty", "Party name cannot be empty.")
                return
            if not code:
                messagebox.showwarning("Empty", "Code cannot be empty.")
                return
            if name in self.party_map:
                messagebox.showwarning("Exists", f"Party '{name}' already exists.")
                return
            self.party_map[name] = code
            add_name_entry.delete(0, "end")
            add_code_entry.delete(0, "end")
            refresh_list()

        ctk.CTkButton(add_frame, text="Add", command=add_party).pack(side="left")

        # === Buttons ===
        btn_frame = ctk.CTkFrame(popup, fg_color="transparent")
        btn_frame.pack(pady=10)

        def save_and_close():
            updated = False
            # Collect all codes before destroying any widget
            updates = []
            for name, code_entry, _ in party_widgets:
                try:
                    new_code = code_entry.get().strip()
                    if new_code and self.party_map.get(name) != new_code:
                        updates.append((name, new_code))
                        updated = True
                except Exception as e:
                    # Ignore destroyed widgets
                    pass

            # Apply updates
            for name, new_code in updates:
                self.party_map[name] = new_code

            self.save_parties_csv()
            if updated:
                self.scan_folder()
            popup.destroy()
            self.status_label.configure(text="✅ Parties saved")

        def reset_to_default():
            default_parties = {
                "Creative": "2",
                "Pranam Maheta": "7",
                "XYZ Designs": "5",
                "Sunrise": "3",
                "Vikas": "9",
                "Kuber": "12"
            }
            self.party_map.clear()
            self.party_map.update(default_parties)
            self.save_parties_csv()
            refresh_list()  # Rebuilds UI safely
            self.status_label.configure(text="🔄 Default parties restored")

        ctk.CTkButton(btn_frame, text="💾 Save & Close", command=save_and_close).pack(side="left", padx=5)
        ctk.CTkButton(btn_frame, text="🔄 Reset", command=reset_to_default).pack(side="left", padx=5)

        # === Finalize Popup ===
        popup.bind("<Return>", lambda e: add_party())  # Enter to add
        search_entry.focus()
        popup.deiconify()  # Show only when fully ready

        # Optional: Keep reference to prevent garbage collection
        self.current_popup = popup

    def save_parties_csv(self):
        """Save party map to CSV"""
        csv_path = codes_dir / "parties.csv"
        try:
            with open(csv_path, mode='w', newline='', encoding='utf-8') as f:
                writer = csv.writer(f)
                writer.writerow(["Party Name", "Code"])
                for name in sorted(self.party_map.keys()):
                    writer.writerow([name, self.party_map[name]])
            logging.info(f"✅ Saved {len(self.party_map)} parties to {csv_path}")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save parties.csv:\n{e}")

    def on_closing(self):
        """Called when user closes the window"""
        self.stop_auto_scan()
        self.save_config()
        self.destroy()

if __name__ == "__main__":
    app = FileRenamerApp()
    app.mainloop()