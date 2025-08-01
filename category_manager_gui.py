#!/usr/bin/env python3
"""
Category Manager GUI
===================

Een GUI voor het beheren van backorder categorieën.
"""

import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext
import json
import os
from category_manager import CategoryManager

class CategoryManagerGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Backorder Categorie Manager")
        self.root.geometry("900x600")
        self.root.minsize(700, 500)
        
        # Category manager instance
        self.category_manager = CategoryManager()
        
        # Variables
        self.category_var = tk.StringVar()
        self.name_var = tk.StringVar()
        self.action_var = tk.StringVar()
        self.color_var = tk.StringVar()
        self.new_item_var = tk.StringVar()
        self.status_var = tk.StringVar()
        
        self.setup_ui()
        self.load_categories()
    
    def setup_ui(self):
        """Setup de gebruikersinterface."""
        # Configureer grid weights
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        
        # Canvas voor scrollbaarheid
        self.canvas = tk.Canvas(self.root, bg='lightgray')
        self.canvas.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Scrollbar
        scrollbar = ttk.Scrollbar(self.root, orient="vertical", command=self.canvas.yview)
        scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
        
        # Configureer canvas
        self.canvas.configure(yscrollcommand=scrollbar.set)
        
        # Hoofdframe binnen canvas
        main_frame = ttk.Frame(self.canvas, padding="20")
        self.canvas.create_window((0, 0), window=main_frame, anchor="nw")
        
        # Configureer main_frame
        main_frame.columnconfigure(1, weight=1)
        main_frame.rowconfigure(1, weight=1)
        
        # Bind events voor scrollbaarheid
        main_frame.bind("<Configure>", lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all")))
        self.canvas.bind("<Configure>", self.on_canvas_configure)
        
        # Bind mousewheel voor scrollen
        self.canvas.bind_all("<MouseWheel>", self.on_mousewheel)
        
        # Header
        self.setup_header(main_frame)
        
        # Categorieën sectie
        self.setup_categories_section(main_frame)
        
        # Item beheer sectie
        self.setup_item_management_section(main_frame)
        
        # Links beheer sectie
        self.setup_links_management_section(main_frame)
        
        # Status sectie
        self.setup_status_section(main_frame)
    
    def setup_header(self, parent):
        """Setup de header."""
        header_frame = ttk.Frame(parent)
        header_frame.grid(row=0, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 20))
        
        title_label = ttk.Label(header_frame, text="📋 Backorder Categorie Manager", 
                               font=("Arial", 18, "bold"))
        title_label.pack()
        
        subtitle_label = ttk.Label(header_frame, 
                                  text="Beheer de categorieën voor backorder artikelen",
                                  font=("Arial", 10), foreground="gray")
        subtitle_label.pack(pady=(5, 0))
    
    def setup_categories_section(self, parent):
        """Setup de categorieën sectie."""
        categories_frame = ttk.LabelFrame(parent, text="📊 Categorieën Overzicht", padding="15")
        categories_frame.grid(row=1, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), padx=(0, 10))
        categories_frame.columnconfigure(0, weight=1)
        categories_frame.rowconfigure(1, weight=1)
        
        # Categorie selectie
        ttk.Label(categories_frame, text="Selecteer categorie:").grid(row=0, column=0, sticky=tk.W, pady=(0, 10))
        
        self.category_var = tk.StringVar(value="category_1")
        category_combo = ttk.Combobox(categories_frame, textvariable=self.category_var, 
                                     values=["category_1", "category_2", "category_3"], 
                                     state="readonly", width=15)
        category_combo.grid(row=0, column=1, sticky=tk.W, pady=(0, 10))
        category_combo.bind("<<ComboboxSelected>>", self.on_category_selected)
        
        # Categorie informatie
        info_frame = ttk.Frame(categories_frame)
        info_frame.grid(row=1, column=0, columnspan=2, sticky=(tk.W, tk.E, tk.N, tk.S))
        info_frame.columnconfigure(1, weight=1)
        
        # Naam
        ttk.Label(info_frame, text="Naam:").grid(row=0, column=0, sticky=tk.W, pady=2)
        self.name_entry = ttk.Entry(info_frame, textvariable=self.name_var, width=40)
        self.name_entry.grid(row=0, column=1, sticky=(tk.W, tk.E), padx=(10, 0), pady=2)
        
        # Beschrijving
        ttk.Label(info_frame, text="Beschrijving:").grid(row=1, column=0, sticky=tk.W, pady=2)
        self.description_text = scrolledtext.ScrolledText(info_frame, height=4, width=40)
        self.description_text.grid(row=1, column=1, sticky=(tk.W, tk.E), padx=(10, 0), pady=2)
        
        # Actie
        ttk.Label(info_frame, text="Actie:").grid(row=2, column=0, sticky=tk.W, pady=2)
        self.action_entry = ttk.Entry(info_frame, textvariable=self.action_var, width=40)
        self.action_entry.grid(row=2, column=1, sticky=(tk.W, tk.E), padx=(10, 0), pady=2)
        
        # Kleur
        ttk.Label(info_frame, text="Kleur (hex):").grid(row=3, column=0, sticky=tk.W, pady=2)
        self.color_entry = ttk.Entry(info_frame, textvariable=self.color_var, width=10)
        self.color_entry.grid(row=3, column=1, sticky=tk.W, padx=(10, 0), pady=2)
        
        # Update knop
        update_button = ttk.Button(info_frame, text="💾 Update Categorie", command=self.update_category)
        update_button.grid(row=4, column=0, columnspan=2, pady=(20, 0))
    
    def setup_item_management_section(self, parent):
        """Setup de item beheer sectie."""
        items_frame = ttk.LabelFrame(parent, text="🔧 Item Beheer", padding="15")
        items_frame.grid(row=1, column=1, sticky=(tk.W, tk.E, tk.N, tk.S))
        items_frame.columnconfigure(0, weight=1)
        items_frame.rowconfigure(1, weight=1)
        
        # Item toevoegen
        add_frame = ttk.Frame(items_frame)
        add_frame.grid(row=0, column=0, sticky=(tk.W, tk.E), pady=(0, 10))
        add_frame.columnconfigure(1, weight=1)
        
        ttk.Label(add_frame, text="Item toevoegen:").grid(row=0, column=0, sticky=tk.W)
        self.new_item_entry = ttk.Entry(add_frame, textvariable=self.new_item_var, width=15)
        self.new_item_entry.grid(row=0, column=1, sticky=(tk.W, tk.E), padx=(10, 10))
        
        add_button = ttk.Button(add_frame, text="➕ Toevoegen", command=self.add_item)
        add_button.grid(row=0, column=2)
        
        # Items lijst
        list_frame = ttk.Frame(items_frame)
        list_frame.grid(row=1, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        list_frame.columnconfigure(0, weight=1)
        list_frame.rowconfigure(0, weight=1)
        
        # Instructie voor items
        instruction_frame = ttk.Frame(list_frame)
        instruction_frame.grid(row=0, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 5))
        
        instruction_label = ttk.Label(instruction_frame, 
                                     text="📋 Items in categorie (klik om te selecteren voor links/verwijderen):",
                                     font=("Arial", 9), foreground="blue")
        instruction_label.pack(side=tk.LEFT)
        
        # Treeview voor items
        self.items_tree = ttk.Treeview(list_frame, columns=("item",), show="tree", height=10)
        self.items_tree.grid(row=1, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Bind click event
        self.items_tree.bind("<ButtonRelease-1>", self.on_tree_item_click)
        
        # Scrollbar
        scrollbar = ttk.Scrollbar(list_frame, orient="vertical", command=self.items_tree.yview)
        scrollbar.grid(row=1, column=1, sticky=(tk.N, tk.S))
        self.items_tree.configure(yscrollcommand=scrollbar.set)
        
        # Geselecteerd item display
        selected_item_frame = ttk.Frame(list_frame)
        selected_item_frame.grid(row=2, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(5, 0))
        selected_item_frame.columnconfigure(1, weight=1)
        
        ttk.Label(selected_item_frame, text="Geselecteerd:").grid(row=0, column=0, sticky=tk.W, pady=2)
        self.selected_item_display = ttk.Label(selected_item_frame, text="Geen item geselecteerd", 
                                             font=("Arial", 9), foreground="gray")
        self.selected_item_display.grid(row=0, column=1, sticky=tk.W, padx=(10, 0), pady=2)
        
        # Knoppen frame
        buttons_frame = ttk.Frame(list_frame)
        buttons_frame.grid(row=3, column=0, columnspan=2, pady=(10, 0))
        
        # Verwijder knop
        remove_button = ttk.Button(buttons_frame, text="🗑️ Verwijder Geselecteerd", command=self.remove_item, 
                                  style="Danger.TButton")
        remove_button.pack(side=tk.LEFT, padx=(0, 10))
        
        # Duidelijke instructie
        instruction_label = ttk.Label(buttons_frame, 
                                     text="← Klik hier om het geselecteerde item te verwijderen",
                                     font=("Arial", 8), foreground="red")
        instruction_label.pack(side=tk.LEFT, padx=(5, 0))
        
        # Statistieken
        self.stats_label = ttk.Label(buttons_frame, text="Totaal items: 0")
        self.stats_label.pack(side=tk.RIGHT)
    
    def setup_links_management_section(self, parent):
        """Setup de links beheer sectie."""
        links_frame = ttk.LabelFrame(parent, text="🔗 Links per Artikel", padding="15")
        links_frame.grid(row=2, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(10, 0))
        links_frame.columnconfigure(1, weight=1)
        
        # Instructie
        instruction_label = ttk.Label(links_frame, 
                                     text="💡 Klik op een artikel in de lijst hierboven om direct links in te vullen",
                                     font=("Arial", 9), foreground="blue")
        instruction_label.grid(row=0, column=0, columnspan=2, sticky=tk.W, pady=(0, 10))
        
        # Geselecteerd artikel display
        selected_frame = ttk.Frame(links_frame)
        selected_frame.grid(row=1, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 10))
        selected_frame.columnconfigure(1, weight=1)
        
        ttk.Label(selected_frame, text="Geselecteerd artikel:").grid(row=0, column=0, sticky=tk.W, pady=2)
        self.selected_item_label = ttk.Label(selected_frame, text="Geen artikel geselecteerd", 
                                           font=("Arial", 10, "bold"), foreground="gray")
        self.selected_item_label.grid(row=0, column=1, sticky=tk.W, padx=(10, 0), pady=2)
        
        # Links sectie
        links_info_frame = ttk.Frame(links_frame)
        links_info_frame.grid(row=2, column=0, columnspan=2, sticky=(tk.W, tk.E))
        links_info_frame.columnconfigure(1, weight=1)
        
        # Fabrikant link
        ttk.Label(links_info_frame, text="Fabrikant link:").grid(row=0, column=0, sticky=tk.W, pady=2)
        self.fabrikant_link_var = tk.StringVar()
        self.fabrikant_link_entry = ttk.Entry(links_info_frame, textvariable=self.fabrikant_link_var, width=50)
        self.fabrikant_link_entry.grid(row=0, column=1, sticky=(tk.W, tk.E), padx=(10, 0), pady=2)
        
        # Externe verkoper link
        ttk.Label(links_info_frame, text="Externe verkoper link:").grid(row=1, column=0, sticky=tk.W, pady=2)
        self.externe_link_var = tk.StringVar()
        self.externe_link_entry = ttk.Entry(links_info_frame, textvariable=self.externe_link_var, width=50)
        self.externe_link_entry.grid(row=1, column=1, sticky=(tk.W, tk.E), padx=(10, 0), pady=2)
        
        # Knoppen frame
        buttons_frame = ttk.Frame(links_frame)
        buttons_frame.grid(row=3, column=0, columnspan=2, pady=(10, 0))
        
        # Save button
        save_links_button = ttk.Button(buttons_frame, text="💾 Links Opslaan", command=self.save_item_links)
        save_links_button.pack(side=tk.LEFT, padx=(0, 10))
        
        # Verwijder knop (duplicaat voor duidelijkheid)
        remove_button_links = ttk.Button(buttons_frame, text="🗑️ Verwijder Artikel", command=self.remove_item)
        remove_button_links.pack(side=tk.LEFT, padx=(0, 10))
        
        # Instructie
        instruction_label = ttk.Label(buttons_frame, 
                                     text="← Of klik hier om het geselecteerde artikel te verwijderen",
                                     font=("Arial", 8), foreground="red")
        instruction_label.pack(side=tk.LEFT, padx=(5, 0))
    
    def setup_status_section(self, parent):
        """Setup de status sectie."""
        status_frame = ttk.Frame(parent)
        status_frame.grid(row=3, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(20, 0))
        
        # Knoppen
        refresh_button = ttk.Button(status_frame, text="🔄 Ververs", command=self.refresh_data)
        refresh_button.pack(side=tk.LEFT, padx=(0, 10))
        
        export_button = ttk.Button(status_frame, text="📤 Export naar Config", command=self.export_to_config)
        export_button.pack(side=tk.LEFT, padx=(0, 10))
        
        # Status label
        self.status_var = tk.StringVar(value="✅ Klaar")
        status_label = ttk.Label(status_frame, textvariable=self.status_var)
        status_label.pack(side=tk.RIGHT)
    
    def on_canvas_configure(self, event):
        """Configureer canvas breedte."""
        self.canvas.itemconfig(self.canvas.find_withtag("all")[0], width=event.width)
    
    def on_mousewheel(self, event):
        """Scroll met muiswiel."""
        self.canvas.yview_scroll(int(-1*(event.delta/120)), "units")
    
    def load_categories(self):
        """Laad categorieën data."""
        self.on_category_selected()
        self.load_all_items()
    
    def on_category_selected(self, event=None):
        """Wanneer een categorie wordt geselecteerd."""
        category_key = self.category_var.get()
        category_num = int(category_key.split("_")[1])
        
        # Laad categorie informatie
        self.name_var.set(self.category_manager.get_category_name(category_num))
        self.action_var.set(self.category_manager.get_category_action(category_num))
        self.color_var.set(self.category_manager.get_category_color(category_num))
        
        # Laad beschrijving
        self.description_text.delete(1.0, tk.END)
        category_data = self.category_manager.categories[category_key]
        self.description_text.insert(1.0, category_data.get("description", ""))
        
        # Laad items
        self.load_items_for_category(category_key)
    
    def load_items_for_category(self, category_key):
        """Laad items voor een categorie."""
        # Wis huidige items
        for item in self.items_tree.get_children():
            self.items_tree.delete(item)
        
        # Voeg nieuwe items toe
        category_num = int(category_key.split("_")[1])
        items = self.category_manager.get_all_items_in_category(category_num)
        for item in sorted(items):
            self.items_tree.insert("", "end", text=item, values=(item,))
        
        # Update statistieken
        self.stats_label.config(text=f"Totaal items: {len(items)}")
    
    def update_category(self):
        """Update categorie informatie."""
        try:
            category_key = self.category_var.get()
            category_num = int(category_key.split("_")[1])
            
            name = self.name_var.get()
            description = self.description_text.get(1.0, tk.END).strip()
            action = self.action_var.get()
            color = self.color_var.get()
            
            if not name:
                messagebox.showerror("Fout", "Naam is verplicht!")
                return
            
            # Update categorie
            self.category_manager.update_category_info(
                category_num, name, description, action, color
            )
            
            self.status_var.set("✅ Categorie bijgewerkt")
            messagebox.showinfo("Succes", "Categorie succesvol bijgewerkt!")
            
        except Exception as e:
            messagebox.showerror("Fout", f"Fout bij bijwerken: {e}")
            self.status_var.set("❌ Fout opgetreden")
    
    def add_item(self):
        """Voeg een item toe aan de huidige categorie."""
        item_no = self.new_item_var.get().strip()
        if not item_no:
            messagebox.showerror("Fout", "Voer een item nummer in!")
            return
        
        category_key = self.category_var.get()
        
        try:
            if self.category_manager.add_item_to_category(item_no, category_key):
                self.new_item_var.set("")  # Wis input
                self.load_items_for_category(category_key)
                self.status_var.set(f"✅ Item {item_no} toegevoegd")
                messagebox.showinfo("Succes", f"Item {item_no} toegevoegd aan {category_key}")
            else:
                messagebox.showwarning("Waarschuwing", f"Item {item_no} staat al in {category_key}")
        except Exception as e:
            messagebox.showerror("Fout", f"Fout bij toevoegen: {e}")
    
    def remove_item(self):
        """Verwijder een geselecteerd item."""
        selection = self.items_tree.selection()
        if not selection:
            messagebox.showwarning("Waarschuwing", "Selecteer eerst een item door erop te klikken!")
            return
        
        item_no = self.items_tree.item(selection[0], "text")
        category_key = self.category_var.get()
        
        if messagebox.askyesno("Bevestig", f"Weet je zeker dat je item {item_no} wilt verwijderen uit {category_key}?"):
            try:
                if self.category_manager.remove_item_from_category(item_no, category_key):
                    # Reset geselecteerd item display
                    self.selected_item_display.config(text="Geen item geselecteerd", foreground="gray")
                    self.selected_item_label.config(text="Geen artikel geselecteerd", foreground="gray")
                    
                    # Herlaad items
                    self.load_items_for_category(category_key)
                    self.status_var.set(f"✅ Item {item_no} verwijderd")
                    messagebox.showinfo("Succes", f"Item {item_no} verwijderd uit {category_key}")
                else:
                    messagebox.showwarning("Waarschuwing", f"Item {item_no} niet gevonden in {category_key}")
            except Exception as e:
                messagebox.showerror("Fout", f"Fout bij verwijderen: {e}")
    
    def on_tree_item_click(self, event=None):
        """Wanneer een artikel in de tree wordt geklikt."""
        selection = self.items_tree.selection()
        if selection:
            item_no = self.items_tree.item(selection[0], "text")
            self.select_item_for_links(item_no)
            # Update geselecteerd item display
            self.selected_item_display.config(text=f"Item {item_no}", foreground="black")
        else:
            # Geen selectie
            self.selected_item_display.config(text="Geen item geselecteerd", foreground="gray")
    
    def select_item_for_links(self, item_no):
        """Selecteer een artikel voor links beheer."""
        # Update geselecteerd artikel label
        self.selected_item_label.config(text=f"Artikel {item_no}", foreground="black")
        
        # Laad bestaande links
        fabrikant_link = self.category_manager.get_item_link(item_no, 'fabrikant')
        self.fabrikant_link_var.set(fabrikant_link)
        
        externe_link = self.category_manager.get_item_link(item_no, 'externe_verkoper')
        self.externe_link_var.set(externe_link)
        
        # Sla geselecteerd item op
        self.selected_item_no = item_no
    
    def on_item_selected(self, event=None):
        """Wanneer een artikel wordt geselecteerd (oude methode - behouden voor compatibiliteit)."""
        item_no = self.item_var.get()
        if item_no:
            self.select_item_for_links(item_no)
    
    def save_item_links(self):
        """Sla links op voor het geselecteerde artikel."""
        if not hasattr(self, 'selected_item_no') or not self.selected_item_no:
            messagebox.showerror("Fout", "Selecteer eerst een artikel door erop te klikken!")
            return
        
        item_no = self.selected_item_no
        fabrikant_link = self.fabrikant_link_var.get().strip()
        externe_link = self.externe_link_var.get().strip()
        
        # Sla links op
        if fabrikant_link:
            self.category_manager.set_item_link(item_no, 'fabrikant', fabrikant_link)
        
        if externe_link:
            self.category_manager.set_item_link(item_no, 'externe_verkoper', externe_link)
        
        self.status_var.set("✅ Links opgeslagen")
        messagebox.showinfo("Succes", f"Links voor artikel {item_no} opgeslagen!")
    
    def load_all_items(self):
        """Laad alle artikelen (niet meer nodig voor combobox, maar behouden voor compatibiliteit)."""
        # Deze functie is niet meer nodig omdat we direct op artikelen klikken
        pass
    
    def refresh_data(self):
        """Ververs alle data."""
        self.category_manager = CategoryManager()
        self.load_categories()
        self.load_all_items()
        self.status_var.set("✅ Data verversd")
    
    def export_to_config(self):
        """Export categorieën naar config formaat."""
        try:
            config_data = self.category_manager.export_to_config_format()
            
            # Toon config data in een popup
            popup = tk.Toplevel(self.root)
            popup.title("Config Export")
            popup.geometry("600x400")
            
            text_widget = scrolledtext.ScrolledText(popup, wrap=tk.WORD)
            text_widget.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
            
            text_widget.insert(1.0, "# Config export voor config.py\n\n")
            text_widget.insert(tk.END, f"CATEGORY_1_ITEMS = {config_data['CATEGORY_1_ITEMS']}\n\n")
            text_widget.insert(tk.END, f"CATEGORY_2_ITEMS = {config_data['CATEGORY_2_ITEMS']}\n\n")
            text_widget.insert(tk.END, f"CATEGORY_3_ITEMS = {config_data['CATEGORY_3_ITEMS']}\n\n")
            text_widget.insert(tk.END, f"CATEGORIES = {config_data['CATEGORIES']}\n")
            
            self.status_var.set("✅ Config export klaar")
            
        except Exception as e:
            messagebox.showerror("Fout", f"Fout bij exporteren: {e}")

def main():
    """Start de GUI."""
    root = tk.Tk()
    
    # Styling
    style = ttk.Style()
    style.theme_use('clam')
    
    # Start GUI
    app = CategoryManagerGUI(root)
    
    # Center window
    root.update_idletasks()
    x = (root.winfo_screenwidth() // 2) - (root.winfo_width() // 2)
    y = (root.winfo_screenheight() // 2) - (root.winfo_height() // 2)
    root.geometry(f"+{x}+{y}")
    
    # Start main loop
    root.mainloop()

if __name__ == "__main__":
    main() 