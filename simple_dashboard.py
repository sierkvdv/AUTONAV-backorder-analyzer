#!/usr/bin/env python3
"""
Navision Backorder Analyzer - Simple Dashboard
=============================================

Een eenvoudig maar modern dashboard voor het analyseren van Navision backorder exports.
Gebruikers kunnen eenvoudig Excel bestanden selecteren en analyseren, inclusief
automatische categorisering en e-mailgeneratie.

Auteur: AUTONAV
Versie: 2.0
Datum: 2024
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import os
import threading
import queue
from datetime import datetime
import sys
import subprocess

# Import het hoofdscript
try:
    from backorder_analyzer import main as run_analysis
    from config import *
except ImportError as e:
    print(f"Fout: Kan hoofdscript niet laden: {e}")
    sys.exit(1)

class SimpleDashboard:
    def __init__(self, root):
        self.root = root
        self.root.title("Navision Backorder Analyzer v2.0")
        self.root.geometry("1000x900")
        self.root.minsize(800, 700)
        self.root.resizable(True, True)
        self.root.configure(bg="lightgray")
        
        # Queue voor thread communicatie
        self.message_queue = queue.Queue()
        
        # Setup UI
        self.setup_ui()
        
        # Start message checker
        self.check_messages()

        # Initial log message
        self.log("🚀 Dashboard gestart. Selecteer een Excel bestand om te beginnen.")
        
    def setup_ui(self):
        """Setup de gebruikersinterface."""

        # Configureer grid weights
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)

        # Canvas met scrollbar
        canvas = tk.Canvas(self.root, bg="lightgray")
        scrollbar = ttk.Scrollbar(self.root, orient="vertical", command=canvas.yview)
        
        # Hoofdframe
        main_frame = ttk.Frame(canvas, padding="20")
        
        # Configureer canvas
        canvas.configure(yscrollcommand=scrollbar.set)
        
        # Plaats alles in grid
        canvas.grid(row=0, column=0, sticky="nsew")
        scrollbar.grid(row=0, column=1, sticky="ns")
        
        # Maak window in canvas met volledige breedte
        canvas_window = canvas.create_window((0, 0), window=main_frame, anchor="nw")
        
        # Update canvas breedte wanneer window resized wordt
        def update_canvas_width(event):
            if event.widget == self.root:
                canvas.itemconfig(canvas_window, width=event.width-20)  # -20 voor scrollbar
        
        self.root.bind('<Configure>', update_canvas_width)
        
        # Configureer main_frame
        main_frame.columnconfigure(0, weight=1)
        main_frame.rowconfigure(6, weight=1)
        
        # Update scrollbar wanneer content verandert
        def configure_scroll(event):
            canvas.configure(scrollregion=canvas.bbox("all"))
        
        main_frame.bind("<Configure>", configure_scroll)
        
        # Header sectie
        self.setup_header(main_frame)
        
        # Bestand selectie sectie
        self.setup_file_section(main_frame)
        
        # Configuratie sectie
        self.setup_config_section(main_frame)
        
        # Categorieën sectie
        self.setup_categories_section(main_frame)
        
        # Log sectie
        self.setup_log_section(main_frame)
        
        # Status en knoppen
        self.setup_status_section(main_frame)
        
        # Category Manager knop wordt toegevoegd aan status sectie
        
    def setup_header(self, parent):
        """Setup de header sectie."""
        header_frame = ttk.Frame(parent)
        header_frame.grid(row=0, column=0, sticky=(tk.W, tk.E), pady=(0, 20))
        
        # Logo en titel container
        logo_title_frame = ttk.Frame(header_frame)
        logo_title_frame.pack()
        
        # QWIC Logo (tekst versie met styling)
        qwic_frame = ttk.Frame(logo_title_frame)
        qwic_frame.pack(side=tk.LEFT, padx=(0, 20))
        
        # QWIC tekst met oranje accent
        qwic_label = tk.Label(qwic_frame, text="QWIC", 
                             font=("Arial Black", 24, "bold"), 
                             fg="#333333", bg="lightgray")
        qwic_label.pack()
        
        # Oranje accent lijn onder QWIC
        accent_line = tk.Frame(qwic_frame, height=3, bg="#FF6B35")  # Oranje kleur
        accent_line.pack(fill=tk.X, pady=(2, 0))
        
        # Navision Analyse Icoon en titel
        navision_frame = ttk.Frame(logo_title_frame)
        navision_frame.pack(side=tk.LEFT)
        
        # Analyse icoon (tekst versie)
        icon_label = tk.Label(navision_frame, text="📈", 
                             font=("Arial", 32), 
                             fg="#0066CC", bg="lightgray")  # Blauwe kleur
        icon_label.pack()
        
        # Titel
        title_label = ttk.Label(navision_frame, text="Navision Backorder Analyzer v2.0",
                               font=("Arial", 18, "bold"))
        title_label.pack()
        
        # Subtitle
        subtitle_label = ttk.Label(navision_frame, text="Analyseer je Navision exports met automatische categorisering en e-mailgeneratie",
                                  font=("Arial", 10), foreground="gray")
        subtitle_label.pack(pady=(5, 0))
        
        # Scheidingslijn
        separator = ttk.Separator(header_frame, orient='horizontal')
        separator.pack(fill=tk.X, pady=(20, 0))
        
    def setup_file_section(self, parent):
        """Setup de bestand selectie sectie."""
        file_frame = ttk.LabelFrame(parent, text="📁 Excel Bestand Selecteren", padding="15")
        file_frame.grid(row=1, column=0, sticky=(tk.W, tk.E), pady=(0, 20))
        file_frame.columnconfigure(1, weight=1)

        # Bestand pad label
        self.file_label = ttk.Label(file_frame, text="❌ Geen bestand geselecteerd",
                                   font=("Arial", 10), foreground="gray")
        self.file_label.grid(row=0, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 15))

        # Knoppen
        button_frame = ttk.Frame(file_frame)
        button_frame.grid(row=1, column=0, columnspan=2, sticky=(tk.W, tk.E))

        self.browse_button = ttk.Button(button_frame, text="📂 Bestand Kiezen",
                                       command=self.browse_file, style="Accent.TButton")
        self.browse_button.pack(side=tk.LEFT, padx=(0, 10))

        self.analyze_button = ttk.Button(button_frame, text="🔍 Analyse Starten",
                                        command=self.start_analysis, state="disabled")
        self.analyze_button.pack(side=tk.LEFT)
        
    def setup_config_section(self, parent):
        """Setup de configuratie sectie."""
        config_frame = ttk.LabelFrame(parent, text="⚙️ Configuratie", padding="15")
        config_frame.grid(row=2, column=0, sticky=(tk.W, tk.E), pady=(0, 20))
        
        # Configuratie opties in een grid
        config_grid = ttk.Frame(config_frame)
        config_grid.pack(fill=tk.X)
        
        # Location Code
        ttk.Label(config_grid, text="📍 Location Code:").grid(row=0, column=0, sticky=tk.W, padx=(0, 10), pady=5)
        self.location_var = tk.StringVar(value=LOCATION_CODE)
        location_entry = ttk.Entry(config_grid, textvariable=self.location_var, width=15)
        location_entry.grid(row=0, column=1, sticky=tk.W, padx=(0, 30), pady=5)

        # Fully Reserved
        ttk.Label(config_grid, text="🔒 Fully Reserved:").grid(row=0, column=2, sticky=tk.W, padx=(0, 10), pady=5)
        self.reserved_var = tk.StringVar(value=FULLY_RESERVED)
        reserved_combo = ttk.Combobox(config_grid, textvariable=self.reserved_var,
                                     values=["Yes", "No"], width=10, state="readonly")
        reserved_combo.grid(row=0, column=3, sticky=tk.W, padx=(0, 30), pady=5)

        # Order Status
        ttk.Label(config_grid, text="📋 Order Status:").grid(row=0, column=4, sticky=tk.W, padx=(0, 10), pady=5)
        self.status_var = tk.StringVar(value=ORDER_STATUS)
        status_entry = ttk.Entry(config_grid, textvariable=self.status_var, width=15)
        status_entry.grid(row=0, column=5, sticky=tk.W, pady=5)
        
    def setup_categories_section(self, parent):
        """Setup de categorieën sectie."""
        categories_frame = ttk.LabelFrame(parent, text="🏷️ Backorder Categorieën", padding="15")
        categories_frame.grid(row=3, column=0, sticky=(tk.W, tk.E), pady=(0, 20))
        
        # Categorieën overzicht
        categories_text = f"""
📋 Categorie 1 - Bestel bij fabrikant (Rood)
   • Artikel is niet meer via QWIC leverbaar
   • Backorder wordt verwijderd
   • Automatische e-mail naar dealer met link naar fabrikant

📋 Categorie 2 - Binnenkort leverbaar (Turquoise)  
   • Backorder blijft staan
   • Geen e-mail (Navision stuurt automatisch bij binnenkomst)

📋 Categorie 3 - Geen voorraadvooruitzicht (Oranje)
   • Artikel komt niet meer of pas over zeer lange tijd
   • Backorder wordt verwijderd
   • Automatische e-mail naar dealer met link naar externe verkoper
        """
        
        categories_label = ttk.Label(categories_frame, text=categories_text, 
                                    font=("Consolas", 9), justify=tk.LEFT)
        categories_label.pack(anchor=tk.W)
        
    def setup_log_section(self, parent):
        """Setup de log sectie."""
        log_frame = ttk.LabelFrame(parent, text="📝 Log Output", padding="15")
        log_frame.grid(row=4, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 10))
        log_frame.columnconfigure(0, weight=1)
        log_frame.rowconfigure(0, weight=1)

        # Log text widget met styling
        self.log_text = scrolledtext.ScrolledText(log_frame, height=8, width=80,
                                                 font=("Consolas", 9), bg="#f8f9fa")
        self.log_text.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
    def setup_status_section(self, parent):
        """Setup de status en actie knoppen sectie."""
        status_frame = ttk.LabelFrame(parent, text="🔧 Acties", padding="15")
        status_frame.grid(row=5, column=0, sticky=(tk.W, tk.E), pady=(0, 10))

        # Status bar
        self.status_var = tk.StringVar(value="✅ Klaar")
        status_bar = ttk.Label(status_frame, textvariable=self.status_var,
                              relief=tk.SUNKEN, anchor=tk.W, padding=(10, 5))
        status_bar.grid(row=0, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 10))

        # Progress bar
        self.progress = ttk.Progressbar(status_frame, mode='indeterminate')
        self.progress.grid(row=1, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 15))

        # Actie knoppen in een grid layout
        button_frame = ttk.Frame(status_frame)
        button_frame.grid(row=2, column=0, columnspan=2, sticky=(tk.W, tk.E))

        # Rij 1 knoppen
        row1_frame = ttk.Frame(button_frame)
        row1_frame.pack(fill=tk.X, pady=(0, 5))

        self.open_output_button = ttk.Button(row1_frame, text="📊 Output Openen",
                                            command=self.open_output, state="normal")
        self.open_output_button.pack(side=tk.LEFT, padx=(0, 10))

        self.open_emails_button = ttk.Button(row1_frame, text="📧 E-mails Openen",
                                            command=self.open_emails, state="normal")
        self.open_emails_button.pack(side=tk.LEFT, padx=(0, 10))

        self.open_log_button = ttk.Button(row1_frame, text="📄 Log Bestand Openen",
                                         command=self.open_log, state="normal")
        self.open_log_button.pack(side=tk.LEFT, padx=(0, 10))

        # Rij 2 knoppen
        row2_frame = ttk.Frame(button_frame)
        row2_frame.pack(fill=tk.X)

        self.clear_log_button = ttk.Button(row2_frame, text="🗑️ Log Wissen",
                                          command=self.clear_log)
        self.clear_log_button.pack(side=tk.LEFT, padx=(0, 10))

        self.open_folder_button = ttk.Button(row2_frame, text="📁 Output Map Openen",
                                            command=self.open_output_folder)
        self.open_folder_button.pack(side=tk.LEFT, padx=(0, 10))

        # Category Manager knop
        self.category_manager_button = ttk.Button(row2_frame, text="📋 Categorie Manager",
                                                command=self.open_category_manager)
        self.category_manager_button.pack(side=tk.LEFT, padx=(0, 10))
        
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
        filename = os.path.basename(file_path)
        self.file_label.config(text=f"✅ Geselecteerd: {filename}",
                              foreground="green")
        self.analyze_button.config(state="normal")
        self.log(f"📁 Bestand geselecteerd: {filename}")
            
    def start_analysis(self):
        """Start de analyse in een aparte thread."""
        if not hasattr(self, 'file_path'):
            messagebox.showerror("❌ Fout", "Selecteer eerst een Excel bestand.")
            return
            
        # Update UI
        self.analyze_button.config(state="disabled")
        self.browse_button.config(state="disabled")
        self.progress.start()
        self.status_var.set("🔄 Analyse bezig...")

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

            # Run analyse met aangepaste file path
            from backorder_analyzer import main
            main(input_file=self.file_path)

            # Success message
            self.message_queue.put(("success", "Analyse succesvol voltooid! 🎉"))
            
        except Exception as e:
            self.message_queue.put(("error", f"Fout tijdens analyse: {str(e)}"))

    def check_messages(self):
        """Check voor berichten van de analyse thread."""
        try:
            while True:
                msg_type, message = self.message_queue.get_nowait()

                if msg_type == "success":
                    self.log(f"✅ {message}")
                    self.status_var.set("✅ Analyse voltooid")
                    self.open_output_button.config(state="normal")
                    self.open_emails_button.config(state="normal")
                    self.open_log_button.config(state="normal")
                    messagebox.showinfo("🎉 Succes", message)
                elif msg_type == "error":
                    self.log(f"❌ {message}")
                    self.status_var.set("❌ Fout opgetreden")
                    messagebox.showerror("❌ Fout", message)

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
        self.log("🗑️ Log gewist.")
            
    def open_output(self):
        """Open het output bestand."""
        output_path = os.path.abspath(OUTPUT_FILE)
        if os.path.exists(output_path):
            try:
                os.startfile(output_path)
            except:
                subprocess.run(['start', output_path], shell=True)
        else:
            messagebox.showerror("❌ Fout", "Output bestand niet gevonden.")
            
    def open_emails(self):
        """Open het e-mail rapport bestand."""
        email_path = OUTPUT_FILE.replace('.xlsx', '_Emails.xlsx')
        email_path = os.path.abspath(email_path)
        if os.path.exists(email_path):
            try:
                os.startfile(email_path)
            except:
                subprocess.run(['start', email_path], shell=True)
        else:
            messagebox.showerror("❌ Fout", "E-mail rapport niet gevonden.")
            
    def open_log(self):
        """Open het log bestand."""
        log_path = os.path.abspath(LOG_FILE)
        if os.path.exists(log_path):
            try:
                os.startfile(log_path)
            except:
                subprocess.run(['start', log_path], shell=True)
        else:
            messagebox.showerror("❌ Fout", "Log bestand niet gevonden.")
            
    def open_output_folder(self):
        """Open de output map."""
        output_dir = os.path.dirname(os.path.abspath(OUTPUT_FILE))
        if os.path.exists(output_dir):
            try:
                os.startfile(output_dir)
            except:
                subprocess.run(['start', output_dir], shell=True)
        else:
            messagebox.showerror("❌ Fout", "Output map niet gevonden.")
    

    
    def open_category_manager(self):
        """Open de Category Manager."""
        try:
            import subprocess
            subprocess.Popen([sys.executable, "category_manager_gui.py"])
        except Exception as e:
            messagebox.showerror("❌ Fout", f"Kan Category Manager niet openen: {e}")

def main():
    """Start het dashboard."""
    root = tk.Tk()

    # Styling
    style = ttk.Style()
    style.theme_use('clam')

    # Custom styling
    style.configure("Accent.TButton", background="#007acc", foreground="white")

    # Start dashboard
    app = SimpleDashboard(root)

    # Center window
    root.update_idletasks()
    x = (root.winfo_screenwidth() // 2) - (root.winfo_width() // 2)
    y = (root.winfo_screenheight() // 2) - (root.winfo_height() // 2)
    root.geometry(f"+{x}+{y}")

    # Start main loop
    root.mainloop()

if __name__ == "__main__":
    main() 