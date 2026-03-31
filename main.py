# =================================================================
# PROJECT: NEURAL-SCAN BIOMETRIC ACCESS TERMINAL (v9.0)
# AUTHOR: OM (Refined by Gemini)
# PURPOSE: B.Tech Final Project - AI Attendance & Analytics
# =================================================================

import time
import sys
import os
import re
import pickle
import cv2
import numpy as np
import pandas as pd
from datetime import datetime
from PIL import Image
from openpyxl import Workbook, load_workbook
import customtkinter as ctk
from tkinter import filedialog
import face_recognition
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from matplotlib.figure import Figure

# --- [ PYINSTALLER SPLASH HANDLER ] ---
try:
    import pyi_splash
except ImportError:
    pyi_splash = None

# --- [ RESOURCE PATH UTILITY ] ---
def resource_path(relative_path):
    """
    Handles absolute paths for both development and PyInstaller environments.
    - Used for AI models and internal assets.
    - NOT used for data persistence (Excel/PKL) to ensure data is writable.
    """
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

# --- [ GLOBAL DESIGN CONFIG ] ---
CYBER_BG      = "#05060F"  # Deepest Midnight
CYBER_SIDEBAR = "#0D0F25"  # Sidebar Navy
CYBER_ACCENT  = "#00F0FF"  # Electric Cyan
CYBER_LABEL   = "#39FF14"  # Matrix Green
CYBER_ERROR   = "#FF2E63"  # Neon Rose
CYBER_DARK    = "#020308"  # Pure Black for Input Fields
CYBER_BORDER  = "#1A1E3D"  # Subtle container borders

FONT_TERMINAL = "Consolas"
SIDE_W        = 340
RADIUS_UI     = 12
BORDER_W      = 2
GAP_S, GAP_M, GAP_L = 8, 16, 28

ctk.set_appearance_mode("Dark")

# =================================================================
# MODULE 1: MAIN APPLICATION CLASS
# =================================================================

class FaceApp(ctk.CTk):
    def __init__(self):
        super().__init__()

        # --- WINDOW CONFIGURATION ---
        self.title("Made with ❤️ by OM")
        self.geometry("1450x880")
        self.configure(fg_color=CYBER_BG)

        # --- DATA & ENGINE INITIALIZATION ---
        # Note: We keep these in the local folder for persistent writing
        self.DB_FILE = "face_data.pkl"
        self.SUBJECTS_FILE = "subjects.pkl"
        
        self.data = self.load_data()
        self.subjects = self.load_subjects()

        # State Variables
        self.cap = None
        self.mode = "Idle"
        self.latest_rgb_frame = None
        self.last_logged_time = {}
        self.chart_canvas = None

        # --- EVENT BINDINGS ---
        # Fix for the "68s" bug - handling both lowercase and uppercase 'S'
        self.bind("<s>", self.save_face_handler)
        self.bind("<S>", self.save_face_handler)

        # --- BASE GRID LAYOUT ---
        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(0, weight=1)

        # Build UI Sections
        self._build_sidebar()
        self._build_tabview()
        self._build_statusbar()

        # Apply Aesthetic Layer
        self.apply_styles()
        
        # Launch Optical Core
        self.start_camera()

    # --- [ UI BUILDERS ] ---

    def _build_sidebar(self):
        """Constructs the primary control column."""
        self.sidebar_frame = ctk.CTkFrame(self, width=SIDE_W, corner_radius=0)
        self.sidebar_frame.grid(row=0, column=0, rowspan=2, sticky="nsew")
        self.sidebar_frame.grid_rowconfigure(10, weight=1) # Spacer to push exit button down

        # 1. Branding
        self.logo_label = ctk.CTkLabel(
            self.sidebar_frame, text="🛡️ NEURAL-SCAN",
            font=(FONT_TERMINAL, 24, "bold"), text_color=CYBER_ACCENT
        )
        self.logo_label.grid(row=0, column=0, padx=GAP_M, pady=(GAP_L, GAP_L))

        # 2. Registration Command
        self.reg_toggle_btn = ctk.CTkButton(
            self.sidebar_frame, text="[ SYNC_NEW_NODE ]",
            height=48, font=(FONT_TERMINAL, 13, "bold"),
            command=self.toggle_registration_form
        )
        self.reg_toggle_btn.grid(row=1, column=0, padx=GAP_M, pady=GAP_S, sticky="ew")

        # 3. Instruction Panel (Dynamic User Feedback)
        self.instruction_frame = ctk.CTkFrame(
            self.sidebar_frame, fg_color=CYBER_DARK, border_width=1, border_color=CYBER_ACCENT
        )
        self.instruction_frame.grid(row=2, column=0, padx=GAP_M, pady=(0, GAP_M), sticky="ew")
        
        self.instruction_label = ctk.CTkLabel(
            self.instruction_frame,
            text=">> SYSTEM_ONLINE\n>> Select protocol to begin.",
            font=(FONT_TERMINAL, 12), text_color=CYBER_LABEL,
            wraplength=SIDE_W - 40, justify="left"
        )
        self.instruction_label.pack(padx=10, pady=10)

        # 4. Hidden Registration Interface (Toggled)
        self.registration_frame = ctk.CTkFrame(self.sidebar_frame, fg_color=CYBER_DARK)
        self.name_entry = ctk.CTkEntry(self.registration_frame, placeholder_text="ID_ALIAS (Letters)", height=40)
        self.name_entry.pack(padx=GAP_M, pady=(GAP_M, GAP_S), fill="x")
        self.roll_entry = ctk.CTkEntry(self.registration_frame, placeholder_text="ROLL_TOKEN (Digits)", height=40)
        self.roll_entry.pack(padx=GAP_M, pady=(0, GAP_S), fill="x")
        self.confirm_reg_btn = ctk.CTkButton(
            self.registration_frame, text="UPLOAD_ENCODING",
            fg_color=CYBER_ACCENT, text_color="black",
            command=self.validate_and_start
        )
        self.confirm_reg_btn.pack(padx=GAP_M, pady=(0, GAP_M), fill="x")

        # 5. Subject Configuration
        self.sub_label = ctk.CTkLabel(self.sidebar_frame, text="PROTOCOL_SUBJECT:", font=(FONT_TERMINAL, 11, "bold"), text_color=CYBER_ACCENT)
        self.sub_label.grid(row=4, column=0, padx=GAP_M, pady=(GAP_L, 4), sticky="w")
        
        self.subject_menu = ctk.CTkComboBox(self.sidebar_frame, values=self.subjects, height=38)
        self.subject_menu.grid(row=5, column=0, padx=GAP_M, pady=(0, GAP_S), sticky="ew")
        if self.subjects: self.subject_menu.set(self.subjects[0])
        
        self.new_sub_btn = ctk.CTkButton(self.sidebar_frame, text="+ ADD_PROTOCOL", height=32, command=self.add_new_subject)
        self.new_sub_btn.grid(row=6, column=0, padx=GAP_M, pady=(0, GAP_M), sticky="ew")

        # 6. Global Action
        self.recog_btn = ctk.CTkButton(
            self.sidebar_frame, text="INITIATE_SCAN", 
            height=64, fg_color="#1A1E3D", border_color=CYBER_LABEL, border_width=2, text_color=CYBER_LABEL,
            font=(FONT_TERMINAL, 16, "bold"), command=self.start_recognition
        )
        self.recog_btn.grid(row=8, column=0, padx=GAP_M, pady=(GAP_L, GAP_S), sticky="ew")

        self.exit_btn = ctk.CTkButton(self.sidebar_frame, text="TERMINATE_SESSION", height=40, fg_color="transparent", border_width=1, border_color="gray", command=self.quit)
        self.exit_btn.grid(row=9, column=0, padx=GAP_M, pady=(0, GAP_L), sticky="ew")

    def _build_tabview(self):
        """Creates the main workspace navigation."""
        self.tabview = ctk.CTkTabview(self, segmented_button_selected_color=CYBER_ACCENT)
        self.tabview.grid(row=0, column=1, padx=GAP_M, pady=(GAP_M, 0), sticky="nsew")
        self.tabview.add("CORE_SCAN")
        self.tabview.add("DATA_VAULT")
        self.tabview.add("ANALYTICS")

        # --- TAB: OPTICAL SCAN ---
        self.scan_tab = self.tabview.tab("CORE_SCAN")
        self.scan_tab.grid_columnconfigure(0, weight=7); self.scan_tab.grid_columnconfigure(1, weight=3)
        self.scan_tab.grid_rowconfigure(0, weight=1)
        
        self.video_label = ctk.CTkLabel(self.scan_tab, text="[ NO_SIGNAL ]", font=(FONT_TERMINAL, 18), text_color=CYBER_ACCENT)
        self.video_label.grid(row=0, column=0, padx=GAP_M, pady=GAP_M, sticky="nsew")
        
        self.log_frame = ctk.CTkFrame(self.scan_tab, fg_color=CYBER_SIDEBAR)
        self.log_frame.grid(row=0, column=1, padx=(0, GAP_M), pady=GAP_M, sticky="nsew")
        
        self.activity_log = ctk.CTkTextbox(self.log_frame, font=(FONT_TERMINAL, 12), fg_color=CYBER_DARK, text_color=CYBER_LABEL)
        self.activity_log.pack(expand=True, fill="both", padx=12, pady=12)

        # --- TAB: DATA VAULT ---
        self.vault_tab = self.tabview.tab("DATA_VAULT")
        self.vault_tab.grid_columnconfigure(0, weight=1); self.vault_tab.grid_rowconfigure(1, weight=1)
        self.scrollable_vault = ctk.CTkScrollableFrame(self.vault_tab, label_text="ID_TOKEN | ALIAS_NAME | ACTIONS", label_text_color=CYBER_ACCENT)
        self.scrollable_vault.grid(row=1, column=0, padx=GAP_L, pady=(0, GAP_L), sticky="nsew")

        # --- TAB: ANALYTICS ---
        self.analytics_tab = self.tabview.tab("ANALYTICS")
        self.analytics_tab.grid_columnconfigure(0, weight=1); self.analytics_tab.grid_rowconfigure(1, weight=1)
        self.analytics_btn = ctk.CTkButton(self.analytics_tab, text="LOAD_ATTENDANCE_LOG (.XLSX)", fg_color=CYBER_SIDEBAR, border_width=1, border_color=CYBER_ACCENT, command=self.browse_attendance_file)
        self.analytics_btn.grid(row=0, column=0, padx=GAP_L, pady=GAP_M, sticky="ew")
        
        self.chart_frame = ctk.CTkFrame(self.analytics_tab, fg_color=CYBER_SIDEBAR)
        self.chart_frame.grid(row=1, column=0, padx=GAP_L, pady=(0, GAP_L), sticky="nsew")

    def _build_statusbar(self):
        self.status_bar = ctk.CTkFrame(self, height=32, corner_radius=0, fg_color=CYBER_DARK)
        self.status_bar.grid(row=1, column=0, columnspan=2, sticky="ew")
        self.status_label = ctk.CTkLabel(self.status_bar, text=">> ENGINE_STANDBY", anchor="w", font=(FONT_TERMINAL, 12), text_color=CYBER_ACCENT)
        self.status_label.pack(side="left", padx=GAP_M)

    # --- [ LOGIC ENGINE ] ---

    def apply_styles(self):
        """Ensures all components follow the Sci-Fi styling guidelines."""
        self.sidebar_frame.configure(fg_color=CYBER_SIDEBAR, border_width=BORDER_W, border_color=CYBER_ACCENT)
        self.registration_frame.configure(border_width=1, border_color=CYBER_ACCENT)
        self.log_frame.configure(border_width=1, border_color=CYBER_BORDER)
        self.refresh_vault()

    def toggle_registration_form(self):
        if self.registration_frame.winfo_viewable():
            self.registration_frame.grid_forget()
            self.reg_toggle_btn.configure(text="[ SYNC_NEW_NODE ]")
        else:
            self.registration_frame.grid(row=3, column=0, sticky="ew", padx=GAP_M, pady=(0, GAP_M))
            self.reg_toggle_btn.configure(text="CANCEL_SYNC")

    def validate_and_start(self):
        """Regex validation for registration fields."""
        name = self.name_entry.get().strip()
        roll = self.roll_entry.get().strip()
        
        if not roll.isdigit():
            self._update_instructions("ERR: Token must be numeric.", CYBER_ERROR)
            return
        if not re.match(r"^[a-zA-Z\s]+$", name):
            self._update_instructions("ERR: Alias must be letters.", CYBER_ERROR)
            return
            
        # CRITICAL: Take focus off input so keyboard 'S' doesn't type into the box
        self.focus_set()
        self.mode = "Register"
        self._update_instructions(f"READY: Press 'S' to capture {name}.", CYBER_ACCENT)

    def save_face_handler(self, event):
        """Captures face encoding and saves to local PKL database."""
        if self.mode == "Register" and self.latest_rgb_frame is not None:
            # Strip 'S' from strings just in case focus was lost
            name = self.name_entry.get().strip().rstrip('sS')
            roll = self.roll_entry.get().strip().rstrip('sS')
            
            self._update_instructions("SCANNING... KEEP STILL", CYBER_ACCENT)
            
            boxes = face_recognition.face_locations(self.latest_rgb_frame)
            encs = face_recognition.face_encodings(self.latest_rgb_frame, boxes)
            
            if encs:
                self.data["names"].append(name)
                self.data["rolls"].append(roll)
                self.data["encodings"].append(encs[0])
                
                with open(self.DB_FILE, "wb") as f:
                    pickle.dump(self.data, f)
                
                self.refresh_vault()
                self.toggle_registration_form()
                self.name_entry.delete(0, 'end'); self.roll_entry.delete(0, 'end')
                self._update_instructions(f"SUCCESS: {name} Registered.", CYBER_LABEL)
                self.mode = "Idle"
            else:
                self._update_instructions("ERR: Face not found. Retrying...", CYBER_ERROR)

    def log_attendance(self, name, roll):
        """Handles Excel operations and log frequency capping."""
        subject = self.subject_menu.get()
        filename = f"{subject}_Attendance.xlsx"
        now = datetime.now()
        
        # 10-second cooldown per student to prevent spam logging
        if roll in self.last_logged_time:
            if (now - self.last_logged_time[roll]).total_seconds() < 10:
                return

        if not os.path.exists(filename):
            wb = Workbook(); ws = wb.active
            ws.append(["Date", "Roll No", "Name", "Time", "Status"])
        else:
            wb = load_workbook(filename); ws = wb.active

        ws.append([now.strftime("%Y-%m-%d"), roll, name, now.strftime("%H:%M:%S"), "Present"])
        wb.save(filename)
        self.last_logged_time[roll] = now
        
        # Update live feed
        self.activity_log.insert("0.0", f"[{now.strftime('%H:%M:%S')}] {name} >> GRANTED\n")

    def refresh_vault(self):
        """Redraws the Student Directory with individual Purge buttons."""
        for widget in self.scrollable_vault.winfo_children(): widget.destroy()
        
        for i, (name, roll) in enumerate(zip(self.data["names"], self.data["rolls"])):
            card = ctk.CTkFrame(self.scrollable_vault, fg_color=CYBER_DARK, height=50, border_width=1, border_color=CYBER_BORDER)
            card.pack(fill="x", pady=4, padx=5)
            
            ctk.CTkLabel(card, text=f"#{roll}", width=80, text_color=CYBER_ACCENT, font=(FONT_TERMINAL, 12, "bold")).pack(side="left", padx=10)
            ctk.CTkLabel(card, text=name, width=200, anchor="w", text_color="white").pack(side="left")
            
            ctk.CTkButton(
                card, text="PURGE", width=60, height=26, fg_color="#922b21", 
                command=lambda idx=i: self.delete_student(idx)
            ).pack(side="right", padx=10)

    def delete_student(self, idx):
        """Permanent removal of biometric node."""
        self.data["names"].pop(idx)
        self.data["rolls"].pop(idx)
        self.data["encodings"].pop(idx)
        with open(self.DB_FILE, "wb") as f:
            pickle.dump(self.data, f)
        self.refresh_vault()
        self._update_instructions("Node purged from database.", CYBER_ERROR)

    def update_frame(self):
        """Core CV2 loop with automated face recognition."""
        ret, frame = self.cap.read()
        if ret:
            rgb_frame = cv2.cvtColor(frame, cv2.COLOR_BGR2RGB)
            self.latest_rgb_frame = rgb_frame
            
            if self.mode == "Recognize":
                # Safety Check: Empty Database protection
                if not self.data["encodings"]:
                    self._update_instructions("ERR: Vault Empty. Sync nodes first.", CYBER_ERROR)
                    self.mode = "Idle"
                else:
                    boxes = face_recognition.face_locations(rgb_frame)
                    encs = face_recognition.face_encodings(rgb_frame, boxes)
                    
                    for box, enc in zip(boxes, encs):
                        dist = face_recognition.face_distance(self.data["encodings"], enc)
                        match_idx = np.argmin(dist)
                        
                        if dist[match_idx] < 0.6: # Confidence Threshold
                            name, roll = self.data["names"][match_idx], self.data["rolls"][match_idx]
                            self.log_attendance(name, roll)
                            
                            t, r, b, l = box
                            cv2.rectangle(frame, (l, t), (r, b), (255, 255, 0), 2)
                            cv2.putText(frame, name, (l, t-10), 1, 0.8, (255, 255, 0), 2)

            # Display Output
            img = Image.fromarray(cv2.cvtColor(frame, cv2.COLOR_BGR2RGB))
            imgtk = ctk.CTkImage(light_image=img, dark_image=img, size=(720, 480))
            self.video_label.configure(image=imgtk, text="")
            
        self.after(10, self.update_frame)

    # --- [ ANALYTICS & DATA ] ---

    def browse_attendance_file(self):
        f = filedialog.askopenfilename(filetypes=[("Excel Data", "*.xlsx")])
        if f:
            df = pd.read_excel(f)
            self.render_analytics(df)

    def render_analytics(self, df):
        """Generates Matplotlib visualization within the CTK frame."""
        if self.chart_canvas: self.chart_canvas.get_tk_widget().destroy()
        
        fig = Figure(figsize=(6, 4), dpi=100); ax = fig.add_subplot(111)
        fig.patch.set_facecolor(CYBER_SIDEBAR); ax.set_facecolor(CYBER_SIDEBAR)
        
        col = "Name" if "Name" in df.columns else df.columns[0]
        counts = df.groupby(col).size()
        
        ax.bar(counts.index.astype(str), counts.values, color=CYBER_ACCENT)
        ax.tick_params(colors='white', labelsize=8)
        ax.set_title("Attendance Frequency Chart", color=CYBER_ACCENT)
        for spine in ax.spines.values(): spine.set_color(CYBER_ACCENT)
        
        canvas = FigureCanvasTkAgg(fig, master=self.chart_frame)
        canvas.draw(); canvas.get_tk_widget().pack(expand=True, fill="both")
        self.chart_canvas = canvas

    # --- [ SYSTEM HELPERS ] ---

    def _update_instructions(self, text, color):
        self.instruction_label.configure(text=f">> {text}", text_color=color)
        self.status_label.configure(text=f">> {text}", text_color=color)

    def load_data(self):
        if os.path.exists(self.DB_FILE):
            with open(self.DB_FILE, "rb") as f: return pickle.load(f)
        return {"names": [], "rolls": [], "encodings": []}

    def load_subjects(self):
        if os.path.exists(self.SUBJECTS_FILE):
            with open(self.SUBJECTS_FILE, "rb") as f: return pickle.load(f)
        return ["Python", "Machine Learning", "Cloud Computing"]

    def save_subjects(self):
        with open(self.SUBJECTS_FILE, "wb") as f: pickle.dump(self.subjects, f)

    def add_new_subject(self):
        dialog = ctk.CTkInputDialog(text="New Protocol ID:", title="PROTOCOL_ENTRY")
        new = dialog.get_input()
        if new and new not in self.subjects:
            self.subjects.append(new); self.save_subjects()
            self.subject_menu.configure(values=self.subjects)

    def start_camera(self):
        if not self.cap: self.cap = cv2.VideoCapture(0)
        self.update_frame()

    def start_recognition(self):
        if self.data["names"]: 
            self.mode = "Recognize"
            self._update_instructions("SCANNER_ACTIVE", CYBER_LABEL)
        else: self._update_instructions("ERR: DATABASE_EMPTY", CYBER_ERROR)

# =================================================================
# MODULE 2: SYSTEM BOOTLOADER (Splash & Progress)
# =================================================================

if __name__ == "__main__":
    # Create Initialization Window
    boot = ctk.CTk()
    boot.title("NEURAL-SCAN BOOT")
    boot.geometry("420x200")
    boot.overrideredirect(True)
    
    # Center on screen
    sw, sh = boot.winfo_screenwidth(), boot.winfo_screenheight()
    boot.geometry(f"+{int(sw/2-210)}+{int(sh/2-100)}")
    boot.configure(fg_color=CYBER_DARK)

    ctk.CTkLabel(boot, text="NEURAL-SCAN TERMINAL v9.0", font=(FONT_TERMINAL, 18, "bold"), text_color=CYBER_ACCENT).pack(pady=(30, 5))
    ctk.CTkLabel(boot, text="DECRYPTING AI WEIGHTS...", font=(FONT_TERMINAL, 11), text_color=CYBER_LABEL).pack()
    
    progress = ctk.CTkProgressBar(boot, width=340, progress_color=CYBER_ACCENT)
    progress.pack(pady=25); progress.set(0)

    def launch_sequence():
        # Simulated loading progress (actual logic loads in background)
        for i in range(1, 11):
            progress.set(i/10); boot.update(); time.sleep(0.12)
        
        # Close the PyInstaller Splash image
        if pyi_splash: pyi_splash.close()
        
        boot.destroy() # Close bootloader
        FaceApp().mainloop() # Start main engine

    boot.after(200, launch_sequence)
    boot.mainloop()