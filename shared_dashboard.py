#!/usr/bin/env python3
"""
AUTONAV Backorder Analyzer - Gedeeld Dashboard
==============================================

Dashboard versie voor gedeeld gebruik in een netwerkomgeving.
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import threading
import queue
import os
import sys
import subprocess
import json
from datetime import datetime
import logging

# Import de analyzer modules
try:
    from backorder_analyzer import main as run_analyzer
    from category_manager import CategoryManager
    from config import INPUT_FILE, OUTPUT_FILE
except ImportError as e:
    print(f"Fout bij importeren modules: {e}")
    print("Zorg dat alle bestanden in dezelfde map staan.")
    sys.exit(1)

class SharedDashboard:
    def __init__(self, root):
        self.root = root
        self.root.title("AUTONAV Backorder Analyzer - Gedeeld Dashboard")
        self.root.geometry("800x700")
        self.root.resizable(True, True)
        
        # Queue voor thread communicatie
        self.log_queue = queue.Queue()
        
        # Setup logging
        self.setup_logging()
        
        # Laad configuratie
        self.category_manager = CategoryManager()
        
        # Setup UI
        self.setup_ui()
        
        # Start log polling
        self.poll_log_queue()
        
        # Toon welkomstbericht
        self.log_message("🚀 AUTONAV Backorder Analyzer gestart")
        self.log_message("📁 Gedeelde versie - Configuratie wordt automatisch gesynchroniseerd")
        
    def setup_logging(self):
        """Setup logging voor gedeeld gebruik."""
        log_file = f"backorder_analyzer_{datetime.now().strftime('%Y%m%d')}.log"
        
        logging.basicConfig(
            level=logging.INFO,
            format='%(asctime)s - %(levelname)s - %(message)s',
            handlers=[
                logging.FileHandler(log_file, encoding='utf-8'),
                logging.StreamHandler()
            ]
        )
        
    def setup_ui(self):
        """Setup de gebruikersinterface."""
        # Hoofdframe
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        main_frame.columnconfigure(0, weight=1)
        main_frame.rowconfigure(2, weight=1)
        
        # Titel
        title_frame = ttk.Frame(main_frame)
        title_frame.grid(row=0, column=0, sticky=(tk.W, tk.E), pady=(0, 20))
        title_frame.columnconfigure(1, weight=1)
        
        ttk.Label(title_frame, text="🚗 AUTONAV Backorder Analyzer", 
                 font=("Arial", 16, "bold")).grid(row=0, column=0, columnspan=2, pady=(0, 5))
        ttk.Label(title_frame, text="Gedeelde versie voor teamgebruik", 
                 font=("Arial", 10)).grid(row=1, column=0, columnspan=2)
        
        # Status indicator
        self.status_var = tk.StringVar(value="🟢 Systeem gereed")
        status_label = ttk.Label(title_frame, textvariable=self.status_var, 
                               font=("Arial", 9, "bold"))
        status_label.grid(row=2, column=0, columnspan=2, pady=(5, 0))
        
        # Configuratie sectie
        self.setup_config_section(main_frame)
        
        # Actie sectie
        self.setup_action_section(main_frame)
        
        # Log sectie
        self.setup_log_section(main_frame)
        
        # Status bar
        self.setup_status_bar(main_frame)
        
    def setup_config_section(self, parent):
        """Setup de configuratie sectie."""
        config_frame = ttk.LabelFrame(parent, text="⚙️ Configuratie", padding="10")
        config_frame.grid(row=1, column=0, sticky=(tk.W, tk.E), pady=(0, 10))
        config_frame.columnconfigure(1, weight=1)
        
        # Bestand selectie
        ttk.Label(config_frame, text="Navision Export:").grid(row=0, column=0, sticky=tk.W, pady=2)
        
        self.file_var = tk.StringVar(value=INPUT_FILE)
        file_entry = ttk.Entry(config_frame, textvariable=self.file_var, width=50)
        file_entry.grid(row=0, column=1, sticky=(tk.W, tk.E), padx=(10, 5), pady=2)
        
        browse_button = ttk.Button(config_frame, text="📁 Bladeren", command=self.browse_file)
        browse_button.grid(row=0, column=2, pady=2)
        
        # Output bestand
        ttk.Label(config_frame, text="Output Bestand:").grid(row=1, column=0, sticky=tk.W, pady=2)
        
        self.output_var = tk.StringVar(value=OUTPUT_FILE)
        output_entry = ttk.Entry(config_frame, textvariable=self.output_var, width=50)
        output_entry.grid(row=1, column=1, sticky=(tk.W, tk.E), padx=(10, 5), pady=2)
        
        # Configuratie info
        info_frame = ttk.Frame(config_frame)
        info_frame.grid(row=2, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(10, 0))
        info_frame.columnconfigure(0, weight=1)
        
        self.config_info_var = tk.StringVar()
        self.update_config_info()
        ttk.Label(info_frame, textvariable=self.config_info_var, 
                 font=("Arial", 9)).grid(row=0, column=0, sticky=tk.W)
        
    def setup_action_section(self, parent):
        """Setup de actie sectie."""
        action_frame = ttk.LabelFrame(parent, text="🎯 Acties", padding="10")
        action_frame.grid(row=2, column=0, sticky=(tk.W, tk.E), pady=(0, 10))
        action_frame.columnconfigure(0, weight=1)
        
        # Knoppen rij 1
        button_frame1 = ttk.Frame(action_frame)
        button_frame1.grid(row=0, column=0, sticky=(tk.W, tk.E), pady=(0, 5))
        button_frame1.columnconfigure(0, weight=1)
        button_frame1.columnconfigure(1, weight=1)
        button_frame1.columnconfigure(2, weight=1)
        
        # Analyse knop
        self.analyze_button = ttk.Button(button_frame1, text="🔍 Analyse Starten", 
                                       command=self.start_analysis, style="Accent.TButton")
        self.analyze_button.grid(row=0, column=0, padx=(0, 5))
        
        # Categorie Manager knop
        category_button = ttk.Button(button_frame1, text="📋 Categorie Manager", 
                                   command=self.open_category_manager)
        category_button.grid(row=0, column=1, padx=5)
        
        # Output openen knop
        output_button = ttk.Button(button_frame1, text="📊 Output Openen", 
                                 command=self.open_output_folder)
        output_button.grid(row=0, column=2, padx=(5, 0))
        
        # Knoppen rij 2
        button_frame2 = ttk.Frame(action_frame)
        button_frame2.grid(row=1, column=0, sticky=(tk.W, tk.E))
        button_frame2.columnconfigure(0, weight=1)
        button_frame2.columnconfigure(1, weight=1)
        button_frame2.columnconfigure(2, weight=1)
        
        # E-mails openen knop
        email_button = ttk.Button(button_frame2, text="📧 E-mails Openen", 
                                command=self.open_email_report)
        email_button.grid(row=0, column=0, padx=(0, 5))
        
        # Configuratie verversen knop
        refresh_button = ttk.Button(button_frame2, text="🔄 Configuratie Verversen", 
                                  command=self.refresh_configuration)
        refresh_button.grid(row=0, column=1, padx=5)
        
        # Help knop
        help_button = ttk.Button(button_frame2, text="❓ Help", 
                               command=self.show_help)
        help_button.grid(row=0, column=2, padx=(5, 0))
        
    def setup_log_section(self, parent):
        """Setup de log sectie."""
        log_frame = ttk.LabelFrame(parent, text="📋 Log Berichten", padding="10")
        log_frame.grid(row=3, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 10))
        log_frame.columnconfigure(0, weight=1)
        log_frame.rowconfigure(0, weight=1)
        
        # Log text widget
        self.log_text = scrolledtext.ScrolledText(log_frame, height=15, width=80, 
                                                font=("Consolas", 9))
        self.log_text.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Log knoppen
        log_button_frame = ttk.Frame(log_frame)
        log_button_frame.grid(row=1, column=0, sticky=(tk.W, tk.E), pady=(5, 0))
        
        clear_button = ttk.Button(log_button_frame, text="🗑️ Log Wissen", 
                                command=self.clear_log)
        clear_button.pack(side=tk.LEFT, padx=(0, 5))
        
        save_log_button = ttk.Button(log_button_frame, text="💾 Log Opslaan", 
                                   command=self.save_log)
        save_log_button.pack(side=tk.LEFT)
        
    def setup_status_bar(self, parent):
        """Setup de status bar."""
        status_frame = ttk.Frame(parent)
        status_frame.grid(row=4, column=0, sticky=(tk.W, tk.E))
        status_frame.columnconfigure(0, weight=1)
        
        # Gebruiker info
        self.user_info_var = tk.StringVar(value="👤 Gebruiker: Onbekend")
        ttk.Label(status_frame, textvariable=self.user_info_var, 
                 font=("Arial", 8)).grid(row=0, column=0, sticky=tk.W)
        
        # Tijd info
        self.time_var = tk.StringVar()
        ttk.Label(status_frame, textvariable=self.time_var, 
                 font=("Arial", 8)).grid(row=0, column=1, sticky=tk.E)
        
        # Update tijd
        self.update_time()
        
    def browse_file(self):
        """Open bestand browser."""
        filename = filedialog.askopenfilename(
            title="Selecteer Navision Export",
            filetypes=[("Excel bestanden", "*.xlsx"), ("Alle bestanden", "*.*")]
        )
        if filename:
            self.file_var.set(filename)
            self.log_message(f"📁 Bestand geselecteerd: {os.path.basename(filename)}")
            
    def start_analysis(self):
        """Start de analyse in een aparte thread."""
        if not self.file_var.get():
            messagebox.showerror("Fout", "Selecteer eerst een Navision export bestand!")
            return
            
        if not os.path.exists(self.file_var.get()):
            messagebox.showerror("Fout", "Het geselecteerde bestand bestaat niet!")
            return
            
        # Disable analyse knop
        self.analyze_button.config(state="disabled")
        self.status_var.set("🟡 Analyse bezig...")
        
        # Start analyse thread
        self.analysis_thread = threading.Thread(target=self.run_analysis_thread)
        self.analysis_thread.daemon = True
        self.analysis_thread.start()
        
    def run_analysis_thread(self):
        """Voer analyse uit in aparte thread."""
        try:
            # Update config met geselecteerde bestanden
            self.update_config_files()
            
            # Start analyse
            self.log_message("🚀 Analyse gestart...")
            run_analyzer()
            
            # Analyse voltooid
            self.log_queue.put(("SUCCESS", "✅ Analyse voltooid!"))
            
        except Exception as e:
            self.log_queue.put(("ERROR", f"❌ Fout tijdens analyse: {e}"))
        finally:
            # Re-enable analyse knop
            self.log_queue.put(("ENABLE_BUTTON", ""))
            
    def update_config_files(self):
        """Update config bestanden met geselecteerde bestanden."""
        try:
            # Update config.py
            config_content = f'''# Navision Backorder Analyzer - Configuratie
# =========================================

# =============================================================================
# BESTANDSINSTELLINGEN
# =============================================================================

# Bestandsnaam van de Navision export (pas dit aan naar jouw bestand)
INPUT_FILE = "{self.file_var.get()}"

# Output bestandsnaam
OUTPUT_FILE = "{self.output_var.get()}"

# =============================================================================
# FILTER CRITERIA
# =============================================================================

# Location Code filter
LOCATION_CODE = "DSV"

# Fully Reserved filter
FULLY_RESERVED = "No"

# Order Status filter
ORDER_STATUS = "Backorder"

# =============================================================================
# BACKORDER CATEGORIEËN
# =============================================================================

# Categorie 1: Bestel bij fabrikant (niet meer leverbaar via QWIC)
CATEGORY_1_ITEMS = [
    # QWIC items die niet meer leverbaar zijn
    '10701', '10705', '10708', '10709', '10710',  # Inactive/Phase-out items
    # Algemene items
    'OIL-5W30-1L', 'OIL-10W40-1L', 'BRAKE-FLUID', 'COOLANT-1L', 'POWER-STEERING',
    'AIR-FILTER', 'OIL-FILTER', 'FUEL-FILTER', 'CABIN-FILTER',
    'BRAKE-PADS-FRONT', 'BRAKE-PADS-REAR', 'BRAKE-DISCS-FRONT', 'BRAKE-DISCS-REAR',
    'BATTERY-60AH', 'BATTERY-70AH', 'BATTERY-80AH',
    'CLUTCH-KIT', 'CLUTCH-DISC', 'CLUTCH-BEARING',
    'GEAR-OIL-75W90', 'GEAR-OIL-80W90',
    'STEERING-PUMP', 'STEERING-BOX', 'STEERING-ROD', 'TIE-ROD-END',
    'EXHAUST-MUFFLER', 'EXHAUST-CAT', 'EXHAUST-PIPE',
    'ALTERNATOR', 'STARTER-MOTOR', 'IGNITION-COIL', 'SPARK-PLUGS',
    'TIRES-205-55-16', 'TIRES-225-45-17', 'WHEEL-BEARING',
    'HEADLIGHT-BULB', 'TAILLIGHT-BULB', 'FOG-LIGHT',
    'CAR-MATS', 'STEERING-WHEEL', 'GEAR-KNOB'
]

# Categorie 2: Binnenkort leverbaar (artikelen die QWIC nog wel kan leveren)
CATEGORY_2_ITEMS = [
    # QWIC items die nog leverbaar zijn
    '10700', '10702', '10703', '10704', '10706', '10707', '10711', '10712', '10713', '10714', '10715', '10716', '10717', '10718', '10719', '10720',
    # Algemene items
    'BRAKE-FLUID-DOT4', 'BRAKE-PADS-FRONT', 'BRAKE-PADS-REAR',
    'OIL-FILTER', 'AIR-FILTER', 'BATTERY-60AH'
]

# Categorie 3: Geen voorraadvooruitzicht (artikelen die niet meer leverbaar zijn)
CATEGORY_3_ITEMS = [
    'EXHAUST-CAT', 'EXHAUST-MUFFLER', 'EXHAUST-PIPE',
    'STEERING-WHEEL', 'GEAR-KNOB', 'CAR-MATS'
]

# Categorie namen en beschrijvingen
CATEGORIES = {{
    1: {{
        'name': 'Bestel bij fabrikant',
        'description': 'Artikel is niet meer via QWIC leverbaar. Backorder wordt verwijderd en dealer krijgt automatische e-mail via Salesforce met link naar fabrikant.',
        'action': 'Verwijder backorder + E-mail naar fabrikant',
        'color': 'FF6B6B'  # Rood
    }},
    2: {{
        'name': 'Binnenkort leverbaar',
        'description': 'Backorder blijft staan en wordt niet gemaild (Navision stuurt automatisch zodra artikel binnen is).',
        'action': 'Behoud backorder',
        'color': '4ECDC4'  # Turquoise
    }},
    3: {{
        'name': 'Geen voorraadvooruitzicht',
        'description': 'Artikel komt niet meer of pas over zeer lange tijd. Backorder wordt verwijderd en dealer krijgt e-mail met link naar externe verkoper.',
        'action': 'Verwijder backorder + E-mail naar externe verkoper',
        'color': 'FFA500'  # Oranje
    }}
}}

# =============================================================================
# E-MAIL INSTELLINGEN
# =============================================================================

# E-mail templates voor elke categorie
EMAIL_TEMPLATES = {{
    1: {{
        'subject': 'Backorder artikel niet meer leverbaar via QWIC - {{item_description}}',
        'body': """
Beste {{customer_name}},

Uw backorder artikel "{{item_description}}" (Artikelnummer: {{item_no}}) is helaas niet meer leverbaar via QWIC.

Wij raden u aan om dit artikel direct bij de fabrikant te bestellen:
🔗 {{manufacturer_link}}

Ordergegevens:
- Ordernummer: {{order_no}}
- Artikel: {{item_description}}
- Aantal: {{quantity}}

Voor vragen kunt u contact opnemen met onze klantenservice.

Met vriendelijke groet,
QWIC Team
"""
    }},
    3: {{
        'subject': 'Backorder artikel niet meer beschikbaar - {{item_description}}',
        'body': """
Beste {{customer_name}},

Uw backorder artikel "{{item_description}}" (Artikelnummer: {{item_no}}) is helaas niet meer beschikbaar via QWIC.

Wij raden u aan om dit artikel bij een externe verkoper te bestellen:
🔗 {{external_seller_link}}

Ordergegevens:
- Ordernummer: {{order_no}}
- Artikel: {{item_description}}
- Aantal: {{quantity}}

Voor vragen kunt u contact opnemen met onze klantenservice.

Met vriendelijke groet,
QWIC Team
"""
    }}
}}
'''
            
            with open('config.py', 'w', encoding='utf-8') as f:
                f.write(config_content)
                
            self.log_message("⚙️ Configuratie bijgewerkt")
            
        except Exception as e:
            self.log_message(f"⚠️ Fout bij bijwerken configuratie: {e}")
            
    def open_category_manager(self):
        """Open de category manager in een apart proces."""
        try:
            subprocess.Popen([sys.executable, 'category_manager_gui.py'])
            self.log_message("📋 Category Manager geopend")
        except Exception as e:
            messagebox.showerror("Fout", f"Kan Category Manager niet openen: {e}")
            
    def open_output_folder(self):
        """Open de output map."""
        try:
            output_dir = os.path.dirname(self.output_var.get())
            if not os.path.exists(output_dir):
                os.makedirs(output_dir)
            
            if sys.platform == "win32":
                os.startfile(output_dir)
            else:
                subprocess.run(["xdg-open", output_dir])
                
            self.log_message("📁 Output map geopend")
        except Exception as e:
            messagebox.showerror("Fout", f"Kan output map niet openen: {e}")
            
    def open_email_report(self):
        """Open het email rapport."""
        try:
            output_dir = os.path.dirname(self.output_var.get())
            email_file = os.path.join(output_dir, 
                                    os.path.basename(self.output_var.get()).replace('.xlsx', '_Emails.xlsx'))
            
            if os.path.exists(email_file):
                if sys.platform == "win32":
                    os.startfile(email_file)
                else:
                    subprocess.run(["xdg-open", email_file])
                self.log_message("📧 Email rapport geopend")
            else:
                messagebox.showwarning("Waarschuwing", "Email rapport bestaat nog niet. Voer eerst een analyse uit.")
        except Exception as e:
            messagebox.showerror("Fout", f"Kan email rapport niet openen: {e}")
            
    def refresh_configuration(self):
        """Ververs de configuratie."""
        try:
            self.category_manager = CategoryManager()
            self.update_config_info()
            self.log_message("🔄 Configuratie verversd")
        except Exception as e:
            self.log_message(f"⚠️ Fout bij verversen configuratie: {e}")
            
    def show_help(self):
        """Toon help informatie."""
        help_text = """
🚗 AUTONAV Backorder Analyzer - Help

📋 HOE TE GEBRUIKEN:
1. Selecteer je Navision export Excel bestand
2. Klik op "🔍 Analyse Starten"
3. Wacht tot de analyse klaar is
4. Bekijk de resultaten in de Output map

📋 CATEGORIEËN BEHEREN:
- Klik op "📋 Categorie Manager" om artikelen toe te voegen
- Gebruik "🔗 Links per Artikel" voor specifieke links
- Wijzigingen worden automatisch gedeeld met het team

📊 RESULTATEN:
- Hoofdanalyse: Backorder_Analyse_vX.xlsx
- Email rapport: Backorder_Analyse_vX_Emails.xlsx

💡 TIPS:
- Zorg dat je schrijfrechten hebt op de netwerkmap
- Maak regelmatig backups van je configuratie
- Test nieuwe artikelen eerst met een kleine dataset

📞 Support: Neem contact op met de systeembeheerder
        """
        
        help_window = tk.Toplevel(self.root)
        help_window.title("Help - AUTONAV Backorder Analyzer")
        help_window.geometry("600x500")
        
        text_widget = scrolledtext.ScrolledText(help_window, wrap=tk.WORD, 
                                              font=("Arial", 10))
        text_widget.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        text_widget.insert(tk.END, help_text)
        text_widget.config(state=tk.DISABLED)
        
    def clear_log(self):
        """Wis de log."""
        self.log_text.delete(1.0, tk.END)
        self.log_message("🗑️ Log gewist")
        
    def save_log(self):
        """Sla de log op."""
        try:
            filename = filedialog.asksaveasfilename(
                title="Log Opslaan",
                defaultextension=".txt",
                filetypes=[("Tekst bestanden", "*.txt"), ("Alle bestanden", "*.*")]
            )
            if filename:
                with open(filename, 'w', encoding='utf-8') as f:
                    f.write(self.log_text.get(1.0, tk.END))
                self.log_message(f"💾 Log opgeslagen: {os.path.basename(filename)}")
        except Exception as e:
            messagebox.showerror("Fout", f"Kan log niet opslaan: {e}")
            
    def update_config_info(self):
        """Update configuratie informatie."""
        try:
            categories = self.category_manager.categories
            total_items = sum(len(cat['items']) for cat in categories.values())
            
            info = f"📊 Configuratie: {len(categories)} categorieën, {total_items} artikelen"
            self.config_info_var.set(info)
        except Exception as e:
            self.config_info_var.set(f"⚠️ Configuratie fout: {e}")
            
    def update_time(self):
        """Update tijd in status bar."""
        current_time = datetime.now().strftime("%d-%m-%Y %H:%M:%S")
        self.time_var.set(f"🕐 {current_time}")
        self.root.after(1000, self.update_time)
        
    def log_message(self, message):
        """Voeg bericht toe aan log."""
        timestamp = datetime.now().strftime("%H:%M:%S")
        log_entry = f"[{timestamp}] {message}\n"
        
        # Voeg toe aan log queue
        self.log_queue.put(("LOG", log_entry))
        
    def poll_log_queue(self):
        """Poll de log queue voor berichten."""
        try:
            while True:
                msg_type, message = self.log_queue.get_nowait()
                
                if msg_type == "LOG":
                    self.log_text.insert(tk.END, message)
                    self.log_text.see(tk.END)
                elif msg_type == "SUCCESS":
                    self.log_message(message)
                    self.status_var.set("🟢 Analyse voltooid")
                elif msg_type == "ERROR":
                    self.log_message(message)
                    self.status_var.set("🔴 Fout opgetreden")
                elif msg_type == "ENABLE_BUTTON":
                    self.analyze_button.config(state="normal")
                    
        except queue.Empty:
            pass
        
        # Poll elke 100ms
        self.root.after(100, self.poll_log_queue)

def main():
    """Start de gedeelde dashboard."""
    root = tk.Tk()
    
    # Styling
    style = ttk.Style()
    style.theme_use('clam')
    
    # Accent button style
    style.configure("Accent.TButton", 
                   background="#0078d4", 
                   foreground="white")
    
    app = SharedDashboard(root)
    
    # Center window
    root.update_idletasks()
    x = (root.winfo_screenwidth() // 2) - (root.winfo_width() // 2)
    y = (root.winfo_screenheight() // 2) - (root.winfo_height() // 2)
    root.geometry(f"+{x}+{y}")
    
    root.mainloop()

if __name__ == "__main__":
    main() 