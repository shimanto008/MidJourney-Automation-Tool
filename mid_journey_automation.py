import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pyautogui
import time
import threading
from docx import Document
import csv
import os
import pyperclip  # ADDED: This is required for instant pasting

class MidJourneyAutomation:
    def __init__(self, root):
        self.root = root
        self.root.title("MidJourney Full Automation Tool - Batch Mode")
        self.root.geometry("700x700")
        self.root.resizable(False, False)
        
        self.prompts = []
        self.is_running = False
        self.is_paused = False
        self.current_index = 0
        self.csv_file = None
        self.save_folder = os.getcwd()
        
        self.setup_ui()
        
    def setup_ui(self):
        # Header
        header_frame = tk.Frame(self.root, bg="#4F46E5", height=80)
        header_frame.pack(fill=tk.X)
        header_frame.pack_propagate(False)
        
        title_label = tk.Label(
            header_frame, 
            text="MidJourney Automation - Batch Mode",
            font=("Arial", 20, "bold"),
            bg="#4F46E5",
            fg="white"
        )
        title_label.pack(pady=25)
        
        # Main container
        main_frame = tk.Frame(self.root, bg="#F3F4F6", padx=20, pady=20)
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # File upload section
        upload_frame = tk.LabelFrame(
            main_frame, 
            text="Step 1: Upload Document",
            font=("Arial", 12, "bold"),
            bg="#F3F4F6",
            padx=15,
            pady=15
        )
        upload_frame.pack(fill=tk.X, pady=(0, 15))
        
        self.upload_btn = tk.Button(
            upload_frame,
            text="📁 Select DOCX File",
            command=self.upload_file,
            bg="#3B82F6",
            fg="white",
            font=("Arial", 11, "bold"),
            padx=20,
            pady=10,
            cursor="hand2",
            relief=tk.FLAT
        )
        self.upload_btn.pack()
        
        self.file_label = tk.Label(
            upload_frame,
            text="No file selected",
            font=("Arial", 10),
            bg="#F3F4F6",
            fg="#6B7280"
        )
        self.file_label.pack(pady=(10, 0))
        
        # Save location
        location_frame = tk.Frame(upload_frame, bg="#F3F4F6")
        location_frame.pack(pady=(10, 0))
        
        tk.Label(
            location_frame,
            text="CSV save location:",
            font=("Arial", 9),
            bg="#F3F4F6",
            fg="#6B7280"
        ).pack(side=tk.LEFT)
        
        self.location_label = tk.Label(
            location_frame,
            text=self.save_folder,
            font=("Arial", 9),
            bg="#F3F4F6",
            fg="#3B82F6",
            cursor="hand2"
        )
        self.location_label.pack(side=tk.LEFT, padx=5)
        self.location_label.bind("<Button-1>", lambda e: self.change_save_location())
        
        tk.Button(
            location_frame,
            text="Change",
            command=self.change_save_location,
            bg="#E5E7EB",
            font=("Arial", 8),
            padx=8,
            pady=2,
            cursor="hand2",
            relief=tk.FLAT
        ).pack(side=tk.LEFT)
        
        # Settings section
        settings_frame = tk.LabelFrame(
            main_frame,
            text="Step 2: Configure Settings",
            font=("Arial", 12, "bold"),
            bg="#F3F4F6",
            padx=15,
            pady=15
        )
        settings_frame.pack(fill=tk.X, pady=(0, 15))
        
        # Mode selection
        mode_frame = tk.Frame(settings_frame, bg="#F3F4F6")
        mode_frame.pack(fill=tk.X, pady=5)
        
        tk.Label(
            mode_frame,
            text="Operation Mode:",
            font=("Arial", 10, "bold"),
            bg="#F3F4F6"
        ).pack(side=tk.LEFT)
        
        self.mode_var = tk.StringVar(value="continuous")
        
        tk.Radiobutton(
            mode_frame,
            text="Continuous (all prompts)",
            variable=self.mode_var,
            value="continuous",
            font=("Arial", 10),
            bg="#F3F4F6",
            command=self.toggle_mode_settings
        ).pack(side=tk.LEFT, padx=10)
        
        tk.Radiobutton(
            mode_frame,
            text="Batch Mode (groups with breaks)",
            variable=self.mode_var,
            value="batch",
            font=("Arial", 10),
            bg="#F3F4F6",
            command=self.toggle_mode_settings
        ).pack(side=tk.LEFT)
        
        # Batch mode settings
        self.batch_frame = tk.Frame(settings_frame, bg="#FEF3C7", relief=tk.SOLID, borderwidth=1)
        self.batch_frame.pack(fill=tk.X, pady=10, padx=5)
        
        batch_inner = tk.Frame(self.batch_frame, bg="#FEF3C7", padx=10, pady=10)
        batch_inner.pack(fill=tk.X)
        
        tk.Label(
            batch_inner,
            text="⚡ Batch Mode Settings:",
            font=("Arial", 10, "bold"),
            bg="#FEF3C7",
            fg="#92400E"
        ).pack(anchor=tk.W)
        
        # Batch size
        batch_size_frame = tk.Frame(batch_inner, bg="#FEF3C7")
        batch_size_frame.pack(fill=tk.X, pady=5)
        
        tk.Label(
            batch_size_frame,
            text="Prompts per batch:",
            font=("Arial", 9),
            bg="#FEF3C7"
        ).pack(side=tk.LEFT)
        
        self.batch_size_var = tk.StringVar(value="10")
        tk.Spinbox(
            batch_size_frame,
            from_=1,
            to=50,
            textvariable=self.batch_size_var,
            width=5,
            font=("Arial", 9)
        ).pack(side=tk.LEFT, padx=10)
        
        tk.Label(
            batch_size_frame,
            text="prompts",
            font=("Arial", 9),
            bg="#FEF3C7"
        ).pack(side=tk.LEFT)
        
        # Break time
        break_time_frame = tk.Frame(batch_inner, bg="#FEF3C7")
        break_time_frame.pack(fill=tk.X, pady=5)
        
        tk.Label(
            break_time_frame,
            text="Break between batches:",
            font=("Arial", 9),
            bg="#FEF3C7"
        ).pack(side=tk.LEFT)
        
        self.break_time_var = tk.StringVar(value="4")
        tk.Spinbox(
            break_time_frame,
            from_=1,
            to=30,
            textvariable=self.break_time_var,
            width=5,
            font=("Arial", 9)
        ).pack(side=tk.LEFT, padx=10)
        
        tk.Label(
            break_time_frame,
            text="minutes",
            font=("Arial", 9),
            bg="#FEF3C7"
        ).pack(side=tk.LEFT)
        
        # Typing speed
        speed_frame = tk.Frame(settings_frame, bg="#F3F4F6")
        speed_frame.pack(fill=tk.X, pady=5)
        
        tk.Label(
            speed_frame,
            text="Typing speed:",
            font=("Arial", 10),
            bg="#F3F4F6"
        ).pack(side=tk.LEFT)
        
        self.speed_var = tk.StringVar(value="instant")
        speed_options = [
            ("Instant", "instant"),
            ("Very Fast", "veryfast"),
            ("Fast", "fast")
        ]
        
        for text, value in speed_options:
            tk.Radiobutton(
                speed_frame,
                text=text,
                variable=self.speed_var,
                value=value,
                font=("Arial", 9),
                bg="#F3F4F6"
            ).pack(side=tk.LEFT, padx=5)
        
        # Delay setting
        delay_frame = tk.Frame(settings_frame, bg="#F3F4F6")
        delay_frame.pack(fill=tk.X, pady=5)
        
        tk.Label(
            delay_frame,
            text="Delay after pressing Enter:",
            font=("Arial", 10),
            bg="#F3F4F6"
        ).pack(side=tk.LEFT)
        
        self.delay_var = tk.StringVar(value="5")
        tk.Spinbox(
            delay_frame,
            from_=1,
            to=30,
            textvariable=self.delay_var,
            width=5,
            font=("Arial", 10)
        ).pack(side=tk.LEFT, padx=10)
        
        tk.Label(
            delay_frame,
            text="seconds",
            font=("Arial", 10),
            bg="#F3F4F6"
        ).pack(side=tk.LEFT)
        
        # Countdown setting
        countdown_frame = tk.Frame(settings_frame, bg="#F3F4F6")
        countdown_frame.pack(fill=tk.X, pady=5)
        
        tk.Label(
            countdown_frame,
            text="Countdown before starting:",
            font=("Arial", 10),
            bg="#F3F4F6"
        ).pack(side=tk.LEFT)
        
        self.countdown_var = tk.StringVar(value="5")
        tk.Spinbox(
            countdown_frame,
            from_=3,
            to=15,
            textvariable=self.countdown_var,
            width=5,
            font=("Arial", 10)
        ).pack(side=tk.LEFT, padx=10)
        
        tk.Label(
            countdown_frame,
            text="seconds",
            font=("Arial", 10),
            bg="#F3F4F6"
        ).pack(side=tk.LEFT)
        
        # Control buttons
        control_frame = tk.LabelFrame(
            main_frame,
            text="Step 3: Run Automation",
            font=("Arial", 12, "bold"),
            bg="#F3F4F6",
            padx=15,
            pady=15
        )
        control_frame.pack(fill=tk.X, pady=(0, 15))
        
        button_frame = tk.Frame(control_frame, bg="#F3F4F6")
        button_frame.pack()
        
        self.start_btn = tk.Button(
            button_frame,
            text="▶ START",
            command=self.start_automation,
            bg="#10B981",
            fg="white",
            font=("Arial", 12, "bold"),
            padx=30,
            pady=12,
            cursor="hand2",
            relief=tk.FLAT,
            state=tk.DISABLED
        )
        self.start_btn.pack(side=tk.LEFT, padx=5)
        
        self.pause_btn = tk.Button(
            button_frame,
            text="⏸ PAUSE",
            command=self.pause_automation,
            bg="#F59E0B",
            fg="white",
            font=("Arial", 12, "bold"),
            padx=30,
            pady=12,
            cursor="hand2",
            relief=tk.FLAT,
            state=tk.DISABLED
        )
        self.pause_btn.pack(side=tk.LEFT, padx=5)
        
        self.stop_btn = tk.Button(
            button_frame,
            text="⏹ STOP",
            command=self.stop_automation,
            bg="#EF4444",
            fg="white",
            font=("Arial", 12, "bold"),
            padx=30,
            pady=12,
            cursor="hand2",
            relief=tk.FLAT,
            state=tk.DISABLED
        )
        self.stop_btn.pack(side=tk.LEFT, padx=5)
        
        # Progress section
        progress_frame = tk.LabelFrame(
            main_frame,
            text="Progress",
            font=("Arial", 12, "bold"),
            bg="#F3F4F6",
            padx=15,
            pady=15
        )
        progress_frame.pack(fill=tk.BOTH, expand=True)
        
        self.progress_label = tk.Label(
            progress_frame,
            text="Ready to start",
            font=("Arial", 11),
            bg="#F3F4F6",
            fg="#374151"
        )
        self.progress_label.pack(pady=(0, 10))
        
        self.progress_bar = ttk.Progressbar(
            progress_frame,
            mode='determinate',
            length=600
        )
        self.progress_bar.pack(pady=(0, 10))
        
        self.status_text = tk.Text(
            progress_frame,
            height=6,
            width=70,
            font=("Courier", 9),
            bg="#FFFFFF",
            fg="#1F2937",
            relief=tk.FLAT,
            padx=10,
            pady=10
        )
        self.status_text.pack()
        
        self.toggle_mode_settings()
        
    def toggle_mode_settings(self):
        if self.mode_var.get() == "batch":
            self.batch_frame.pack(fill=tk.X, pady=10, padx=5, after=self.mode_var.master)
        else:
            self.batch_frame.pack_forget()
            
    def change_save_location(self):
        folder = filedialog.askdirectory(
            title="Select Folder to Save CSV Files",
            initialdir=self.save_folder
        )
        if folder:
            self.save_folder = folder
            self.location_label.config(text=folder)
            self.log_status(f"✓ CSV save location changed to: {folder}")
            
    def upload_file(self):
        file_path = filedialog.askopenfilename(
            title="Select DOCX File",
            filetypes=[("Word Documents", "*.docx")]
        )
        
        if not file_path:
            return
            
        try:
            doc = Document(file_path)
            self.prompts = []
            
            for para in doc.paragraphs:
                text = para.text.strip()
                if text:
                    self.prompts.append(text)
            
            if not self.prompts:
                messagebox.showerror("Error", "No prompts found in the document!")
                return
            
            csv_filename = f"prompts_{int(time.time())}.csv"
            self.csv_file = os.path.join(self.save_folder, csv_filename)
            
            with open(self.csv_file, 'w', newline='', encoding='utf-8') as f:
                writer = csv.writer(f)
                for prompt in self.prompts:
                    writer.writerow([prompt])
            
            filename = os.path.basename(file_path)
            self.file_label.config(
                text=f"✓ {filename} ({len(self.prompts)} prompts loaded)",
                fg="#10B981"
            )
            self.start_btn.config(state=tk.NORMAL)
            self.log_status(f"✓ Loaded {len(self.prompts)} prompts from {filename}")
            self.log_status(f"✓ CSV file created: {csv_filename}")
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to read file:\n{str(e)}")
            
    def log_status(self, message):
        self.status_text.insert(tk.END, f"{message}\n")
        self.status_text.see(tk.END)
        self.status_text.update()
        
    def start_automation(self):
        if not self.prompts:
            messagebox.showwarning("Warning", "Please upload a DOCX file first!")
            return
        
        if self.is_paused:
            self.is_paused = False
            self.is_running = True
            self.pause_btn.config(state=tk.NORMAL)
            self.log_status("▶ Resuming automation...")
            threading.Thread(target=self.run_automation, daemon=True).start()
            return
        
        self.current_index = 0
        self.is_running = True
        self.is_paused = False
        
        self.start_btn.config(state=tk.DISABLED)
        self.pause_btn.config(state=tk.NORMAL)
        self.stop_btn.config(state=tk.NORMAL)
        self.upload_btn.config(state=tk.DISABLED)
        
        countdown = int(self.countdown_var.get())
        mode = self.mode_var.get()
        
        self.log_status(f"\n{'='*60}")
        if mode == "batch":
            batch_size = int(self.batch_size_var.get())
            break_time = int(self.break_time_var.get())
            self.log_status(f"🚀 Starting BATCH MODE in {countdown} seconds...")
            self.log_status(f"⚡ {batch_size} prompts per batch, {break_time} min breaks")
        else:
            self.log_status(f"🚀 Starting CONTINUOUS MODE in {countdown} seconds...")
        self.log_status("⚠ SWITCH TO MIDJOURNEY WINDOW NOW!")
        self.log_status("⚠ CLICK IN THE TEXT INPUT BOX!")
        self.log_status(f"{'='*60}\n")
        
        for i in range(countdown, 0, -1):
            self.progress_label.config(text=f"Starting in {i} seconds... Switch to MidJourney!")
            time.sleep(1)
            
        threading.Thread(target=self.run_automation, daemon=True).start()
        
    def run_automation(self):
        delay = int(self.delay_var.get())
        total = len(self.prompts)
        speed = self.speed_var.get()
        mode = self.mode_var.get()
        
        speed_map = {
            "instant": 0,
            "veryfast": 0.001,
            "fast": 0.005
        }
        interval = speed_map.get(speed, 0)
        
        if mode == "batch":
            batch_size = int(self.batch_size_var.get())
            break_time = int(self.break_time_var.get()) * 60  # Convert to seconds
            batch_number = 1
            
        while self.current_index < total and self.is_running:
            if self.is_paused:
                time.sleep(0.1)
                continue
            
            # Check if we need to take a break (batch mode)
            if mode == "batch" and self.current_index > 0 and self.current_index % batch_size == 0:
                self.log_status(f"\n{'='*60}")
                self.log_status(f"☕ BREAK TIME! Batch {batch_number} completed")
                self.log_status(f"⏳ Waiting {int(break_time/60)} minutes before next batch...")
                self.log_status(f"{'='*60}\n")
                
                # Countdown for break
                for remaining in range(break_time, 0, -1):
                    if not self.is_running or self.is_paused:
                        break
                    mins, secs = divmod(remaining, 60)
                    self.progress_label.config(
                        text=f"Break time: {mins:02d}:{secs:02d} remaining | Completed: {self.current_index}/{total}"
                    )
                    time.sleep(1)
                
                batch_number += 1
                self.log_status(f"▶ Starting Batch {batch_number}...\n")
                
            prompt = self.prompts[self.current_index]
            
            self.progress_label.config(
                text=f"Processing prompt {self.current_index + 1} of {total}"
            )
            self.progress_bar['value'] = ((self.current_index + 1) / total) * 100
            
            self.log_status(f"[{self.current_index + 1}/{total}] Pasting: {prompt[:50]}...")
            
            try:
                # ==========================================
                # THE FIX IS HERE
                # ==========================================
                if speed == "instant":
                    pyperclip.copy(prompt)
                    time.sleep(0.1) # Tiny buffer to ensure system clipboard caught up
                    pyautogui.hotkey('ctrl', 'v')
                else:
                    pyautogui.write(prompt, interval=interval)
                # ==========================================
                
                time.sleep(0.1)
                pyautogui.press('enter')
                self.log_status(f"✓ Prompt {self.current_index + 1} submitted!")
                
                self.current_index += 1
                
                if self.current_index < total:
                    self.log_status(f"⏳ Waiting {delay} seconds...\n")
                    time.sleep(delay)
                    
            except Exception as e:
                self.log_status(f"✗ Error: {str(e)}")
                break
        
        if self.current_index >= total and self.is_running:
            self.progress_label.config(text="✅ All prompts completed!")
            self.log_status(f"\n{'='*60}")
            self.log_status("🎉 AUTOMATION COMPLETED!")
            self.log_status(f"✓ Processed {total} prompts")
            if mode == "batch":
                total_batches = (total + int(self.batch_size_var.get()) - 1) // int(self.batch_size_var.get())
                self.log_status(f"✓ Completed {total_batches} batches")
            self.log_status(f"{'='*60}\n")
            self.stop_automation()
            
    def pause_automation(self):
        if self.is_running:
            self.is_paused = True
            self.is_running = False
            self.pause_btn.config(text="▶ RESUME")
            self.start_btn.config(state=tk.NORMAL)
            self.log_status("⏸ Automation paused")
            self.progress_label.config(text="Paused - Click RESUME to continue")
        else:
            self.is_paused = False
            self.is_running = True
            self.pause_btn.config(text="⏸ PAUSE")
            self.start_btn.config(state=tk.DISABLED)
            self.log_status("▶ Automation resumed")
            
    def stop_automation(self):
        self.is_running = False
        self.is_paused = False
        self.current_index = 0
        
        self.start_btn.config(state=tk.NORMAL)
        self.pause_btn.config(state=tk.DISABLED, text="⏸ PAUSE")
        self.stop_btn.config(state=tk.DISABLED)
        self.upload_btn.config(state=tk.NORMAL)
        self.progress_bar['value'] = 0
        self.progress_label.config(text="Stopped")
        self.log_status("⏹ Automation stopped\n")

if __name__ == "__main__":
    pyautogui.FAILSAFE = True
    pyautogui.PAUSE = 0.05
    
    root = tk.Tk()
    app = MidJourneyAutomation(root)
    root.mainloop()