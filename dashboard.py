#!/usr/bin/env python3
"""
Navision Backorder Analyzer - Dashboard
======================================

Een gebruiksvriendelijk dashboard voor het analyseren van Navision backorder exports.
Gebruikers kunnen eenvoudig Excel bestanden slepen en neerzetten.

Auteur: AUTONAV
Versie: 1.0
Datum: 2024
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import os
import threading
import queue
from datetime import datetime
import sys

# Import het hoofdscript
try:
    from backorder_analyzer import main as run_analysis
    from config import *
except ImportError as e:
    messagebox.showerror("Fout", f"Kan hoofdscript niet laden: {e}")
    sys.exit(1)

class BackorderAnalyzerDashboard:
    def __init__(self, root):
        self.root = root
        self.root.title("Navision Backorder Analyzer - Dashboard")
        self.root.geometry("800x600")
        self.root.minsize(600, 400)
        
        # Queue voor thread communicatie
        self.message_queue = queue.Queue()
        
        # Setup UI
        self.setup_ui()
        self.setup_drag_drop()
        
        # Start message checker
        self.check_messages()
    
    def setup_ui(self):
        """Setup de gebruikersinterface."""
        
        # Hoofdframe
        main_frame = ttk.Frame(self.root, padding="20")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Configureer grid weights
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)
        main_frame.rowconfigure(3, weight=1)
        
        # Titel
        title_label = ttk.Label(main_frame, text="Navision Backorder Analyzer", 
                               font=("Arial", 16, "bold"))
        title_label.grid(row=0, column=0, columnspan=3, pady=(0, 20))
        
        # Bestand selectie sectie
        file_frame = ttk.LabelFrame(main_frame, text="Excel Bestand Selecteren", padding="10")
        file_frame.grid(row=1, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 20))
        file_frame.columnconfigure(1, weight=1)
        
        # Bestand pad label
        self.file_label = ttk.Label(file_frame, text="Geen bestand geselecteerd", 
                                   foreground="gray")
        self.file_label.grid(row=0, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 10))
        
        # Knoppen
        button_frame = ttk.Frame(file_frame)
        button_frame.grid(row=1, column=0, columnspan=2, sticky=(tk.W, tk.E))
        
        self.browse_button = ttk.Button(button_frame, text="Bestand Kiezen", 
                                       command=self.browse_file)
        self.browse_button.pack(side=tk.LEFT, padx=(0, 10))
        
        self.analyze_button = ttk.Button(button_frame, text="Analyse Starten", 
                                        command=self.start_analysis, state="disabled")
        self.analyze_button.pack(side=tk.LEFT)
        
        # Configuratie sectie
        config_frame = ttk.LabelFrame(main_frame, text="Configuratie", padding="10")
        config_frame.grid(row=2, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 20))
        
        # Configuratie opties
        ttk.Label(config_frame, text="Location Code:").grid(row=0, column=0, sticky=tk.W, padx=(0, 10))
        self.location_var = tk.StringVar(value=LOCATION_CODE)
        location_entry = ttk.Entry(config_frame, textvariable=self.location_var, width=15)
        location_entry.grid(row=0, column=1, sticky=tk.W, padx=(0, 20))
        
        ttk.Label(config_frame, text="Fully Reserved:").grid(row=0, column=2, sticky=tk.W, padx=(0, 10))
        self.reserved_var = tk.StringVar(value=FULLY_RESERVED)
        reserved_combo = ttk.Combobox(config_frame, textvariable=self.reserved_var, 
                                     values=["Yes", "No"], width=10, state="readonly")
        reserved_combo.grid(row=0, column=3, sticky=tk.W, padx=(0, 20))
        
        ttk.Label(config_frame, text="Order Status:").grid(row=0, column=4, sticky=tk.W, padx=(0, 10))
        self.status_var = tk.StringVar(value=ORDER_STATUS)
        status_entry = ttk.Entry(config_frame, textvariable=self.status_var, width=15)
        status_entry.grid(row=0, column=5, sticky=tk.W)
        
        # Log sectie
        log_frame = ttk.LabelFrame(main_frame, text="Log Output", padding="10")
        log_frame.grid(row=3, column=0, columnspan=3, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 20))
        log_frame.columnconfigure(0, weight=1)
        log_frame.rowconfigure(0, weight=1)
        
        # Log text widget
        self.log_text = scrolledtext.ScrolledText(log_frame, height=15, width=80)
        self.log_text.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Status bar
        self.status_var = tk.StringVar(value="Klaar")
        status_bar = ttk.Label(main_frame, textvariable=self.status_var, 
                              relief=tk.SUNKEN, anchor=tk.W)
        status_bar.grid(row=4, column=0, columnspan=3, sticky=(tk.W, tk.E))
        
        # Progress bar
        self.progress = ttk.Progressbar(main_frame, mode='indeterminate')
        self.progress.grid(row=5, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(5, 0))
        
        # Actie knoppen
        action_frame = ttk.Frame(main_frame)
        action_frame.grid(row=6, column=0, columnspan=3, pady=(10, 0))
        
        self.open_output_button = ttk.Button(action_frame, text="Output Openen", 
                                            command=self.open_output, state="disabled")
        self.open_output_button.pack(side=tk.LEFT, padx=(0, 10))
        
        self.open_log_button = ttk.Button(action_frame, text="Log Bestand Openen", 
                                         command=self.open_log, state="disabled")
        self.open_log_button.pack(side=tk.LEFT, padx=(0, 10))
        
        self.clear_log_button = ttk.Button(action_frame, text="Log Wissen", 
                                          command=self.clear_log)
        self.clear_log_button.pack(side=tk.LEFT)
        
        # Initial log message
        self.log("Dashboard gestart. Sleep een Excel bestand hierheen of klik op 'Bestand Kiezen'.")
    
    def setup_drag_drop(self):
        """Setup drag and drop functionaliteit."""
        self.root.drop_target_register('DND_Files')
        self.root.dnd_bind('<<Drop>>', self.on_drop)
    
    def on_drop(self, event):
        """Handle file drop."""
        files = event.data
        if files:
            file_path = files[0]
            if file_path.lower().endswith('.xlsx'):
                self.set_file_path(file_path)
            else:
                messagebox.showwarning("Ongeldig bestand", 
                                     "Alleen .xlsx bestanden worden ondersteund.")
    
    def browse_file(self):
        """Open file browser."""
        file_path = filedialog.askopenfilename(
            title="Selecteer Navision Export Excel bestand",
            filetypes=[("Excel bestanden", "*.xlsx"), ("Alle bestanden", "*.*")]
        )
        if file_path:
            self.set_file_path(file_path)
    
    def set_file_path(self, file_path):
        """Set het geselecteerde bestand."""
        self.file_path = file_path
        self.file_label.config(text=f"Geselecteerd: {os.path.basename(file_path)}", 
                              foreground="black")
        self.analyze_button.config(state="normal")
        self.log(f"Bestand geselecteerd: {file_path}")
    
    def start_analysis(self):
        """Start de analyse in een aparte thread."""
        if not hasattr(self, 'file_path'):
            messagebox.showerror("Fout", "Selecteer eerst een Excel bestand.")
            return
        
        # Update UI
        self.analyze_button.config(state="disabled")
        self.browse_button.config(state="disabled")
        self.progress.start()
        self.status_var.set("Analyse bezig...")
        
        # Start analyse thread
        thread = threading.Thread(target=self.run_analysis_thread, daemon=True)
        thread.start()
    
    def run_analysis_thread(self):
        """Run de analyse in een aparte thread."""
        try:
            # Update configuratie
            global LOCATION_CODE, FULLY_RESERVED, ORDER_STATUS
            LOCATION_CODE = self.location_var.get()
            FULLY_RESERVED = self.reserved_var.get()
            ORDER_STATUS = self.status_var.get()
            
            # Update input file
            global INPUT_FILE
            INPUT_FILE = self.file_path
            
            # Run analyse
            run_analysis()
            
            # Success message
            self.message_queue.put(("success", "Analyse succesvol voltooid!"))
            
        except Exception as e:
            self.message_queue.put(("error", f"Fout tijdens analyse: {str(e)}"))
    
    def check_messages(self):
        """Check voor berichten van de analyse thread."""
        try:
            while True:
                msg_type, message = self.message_queue.get_nowait()
                
                if msg_type == "success":
                    self.log(f"✓ {message}")
                    self.status_var.set("Analyse voltooid")
                    self.open_output_button.config(state="normal")
                    self.open_log_button.config(state="normal")
                    messagebox.showinfo("Succes", message)
                elif msg_type == "error":
                    self.log(f"✗ {message}")
                    self.status_var.set("Fout opgetreden")
                    messagebox.showerror("Fout", message)
                
                # Reset UI
                self.analyze_button.config(state="normal")
                self.browse_button.config(state="normal")
                self.progress.stop()
                
        except queue.Empty:
            pass
        
        # Check elke 100ms
        self.root.after(100, self.check_messages)
    
    def log(self, message):
        """Voeg bericht toe aan log."""
        timestamp = datetime.now().strftime("%H:%M:%S")
        self.log_text.insert(tk.END, f"[{timestamp}] {message}\n")
        self.log_text.see(tk.END)
    
    def clear_log(self):
        """Wist de log."""
        self.log_text.delete(1.0, tk.END)
        self.log("Log gewist.")
    
    def open_output(self):
        """Open het output bestand."""
        output_path = os.path.abspath(OUTPUT_FILE)
        if os.path.exists(output_path):
            os.startfile(output_path)
        else:
            messagebox.showerror("Fout", "Output bestand niet gevonden.")
    
    def open_log(self):
        """Open het log bestand."""
        log_path = os.path.abspath(LOG_FILE)
        if os.path.exists(log_path):
            os.startfile(log_path)
        else:
            messagebox.showerror("Fout", "Log bestand niet gevonden.")

def main():
    """Start het dashboard."""
    root = tk.Tk()
    
    # Styling
    style = ttk.Style()
    style.theme_use('clam')
    
    # Start dashboard
    app = BackorderAnalyzerDashboard(root)
    
    # Center window
    root.update_idletasks()
    x = (root.winfo_screenwidth() // 2) - (root.winfo_width() // 2)
    y = (root.winfo_screenheight() // 2) - (root.winfo_height() // 2)
    root.geometry(f"+{x}+{y}")
    
    # Start main loop
    root.mainloop()

if __name__ == "__main__":
    main()