#!/usr/bin/env python3
"""
Category Manager
===============

Beheer de backorder categorieën dynamisch.
"""

import json
import os
from datetime import datetime

class CategoryManager:
    def __init__(self, config_file="category_config.json"):
        self.config_file = config_file
        self.categories = self.load_categories()
        self.item_links = self.load_item_links()
    
    def load_categories(self):
        """Laad categorieën uit config bestand."""
        if os.path.exists(self.config_file):
            try:
                with open(self.config_file, 'r', encoding='utf-8') as f:
                    return json.load(f)
            except:
                pass
        
        # Default categorieën
        return {
            "category_1": {
                "name": "Bestel bij fabrikant",
                "description": "Artikel is niet meer via QWIC leverbaar. Backorder wordt verwijderd en dealer krijgt automatische e-mail via Salesforce met link naar fabrikant.",
                "action": "Verwijder backorder + E-mail naar fabrikant",
                "color": "FF6B6B",
                "items": [10701, 10705, 10708, 10709, 10710]
            },
            "category_2": {
                "name": "Binnenkort leverbaar",
                "description": "Backorder blijft staan en wordt niet gemaild (Navision stuurt automatisch zodra artikel binnen is).",
                "action": "Behoud backorder",
                "color": "4ECDC4",
                "items": [10700, 10702, 10703, 10704, 10706, 10707, 10711, 10712, 10713, 10714, 10715, 10718, 10719, 10720]
            },
            "category_3": {
                "name": "Geen voorraadvooruitzicht",
                "description": "Artikel komt niet meer of pas over zeer lange tijd. Backorder wordt verwijderd en dealer krijgt e-mail met link naar externe verkoper.",
                "action": "Verwijder backorder + E-mail naar externe verkoper",
                "color": "FFA500",
                "items": [10716, 10717]
            }
        }
    
    def load_item_links(self):
        """Laad specifieke links per artikel."""
        links_file = "item_links.json"
        if os.path.exists(links_file):
            try:
                with open(links_file, 'r', encoding='utf-8') as f:
                    return json.load(f)
            except:
                pass
        
        # Default links voor test artikelen
        return {
            "10701": {
                "fabrikant": "https://www.shimano.com/nl/products/cycling/",
                "externe_verkoper": "https://www.bike-components.nl/nl/Shimano/"
            },
            "10705": {
                "fabrikant": "https://www.sram.com/nl/",
                "externe_verkoper": "https://www.bike-components.nl/nl/SRAM/"
            },
            "10708": {
                "fabrikant": "https://www.campagnolo.com/nl/",
                "externe_verkoper": "https://www.bike-components.nl/nl/Campagnolo/"
            },
            "10709": {
                "fabrikant": "https://www.michelin.com/nl/",
                "externe_verkoper": "https://www.bike-components.nl/nl/Michelin/"
            },
            "10710": {
                "fabrikant": "https://www.continental-tires.com/nl/",
                "externe_verkoper": "https://www.bike-components.nl/nl/Continental/"
            },
            "10716": {
                "fabrikant": "https://www.bosch-ebike.com/nl/",
                "externe_verkoper": "https://www.bike-components.nl/nl/Bosch/"
            },
            "10717": {
                "fabrikant": "https://www.brose-ebike.com/nl/",
                "externe_verkoper": "https://www.bike-components.nl/nl/Brose/"
            }
        }
    
    def save_categories(self):
        """Sla categorieën op in config bestand."""
        with open(self.config_file, 'w', encoding='utf-8') as f:
            json.dump(self.categories, f, indent=2, ensure_ascii=False)
    
    def save_item_links(self):
        """Sla specifieke links op in config bestand."""
        with open("item_links.json", 'w', encoding='utf-8') as f:
            json.dump(self.item_links, f, indent=2, ensure_ascii=False)
    
    def add_item_to_category(self, item_no, category_key):
        """Voeg een item toe aan een categorie."""
        if category_key in self.categories:
            if item_no not in self.categories[category_key]["items"]:
                self.categories[category_key]["items"].append(item_no)
                self.save_categories()
                return True
        return False
    
    def remove_item_from_category(self, item_no, category_key):
        """Verwijder een item uit een categorie."""
        if category_key in self.categories:
            if item_no in self.categories[category_key]["items"]:
                self.categories[category_key]["items"].remove(item_no)
                self.save_categories()
                return True
        return False
    
    def get_category_for_item(self, item_no):
        """Bepaal de categorie voor een item."""
        for category_key, category_data in self.categories.items():
            if item_no in category_data["items"]:
                return int(category_key.split("_")[1])
        return 2  # Default naar categorie 2
    
    def get_category_name(self, category_number):
        """Krijg de naam van een categorie."""
        category_key = f"category_{category_number}"
        if category_key in self.categories:
            return self.categories[category_key]["name"]
        return "Onbekend"
    
    def get_category_action(self, category_number):
        """Krijg de actie van een categorie."""
        category_key = f"category_{category_number}"
        if category_key in self.categories:
            return self.categories[category_key]["action"]
        return "Onbekend"
    
    def get_category_color(self, category_number):
        """Krijg de kleur van een categorie."""
        category_key = f"category_{category_number}"
        if category_key in self.categories:
            return self.categories[category_key]["color"]
        return "FFFFFF"
    
    def get_item_link(self, item_no, link_type="fabrikant"):
        """Krijg de specifieke link voor een artikel."""
        item_str = str(item_no)
        if item_str in self.item_links:
            return self.item_links[item_str].get(link_type, "")
        return ""
    
    def set_item_link(self, item_no, link_type, url):
        """Stel een specifieke link in voor een artikel."""
        item_str = str(item_no)
        if item_str not in self.item_links:
            self.item_links[item_str] = {}
        self.item_links[item_str][link_type] = url
        self.save_item_links()
    
    def get_all_items_in_category(self, category_number):
        """Krijg alle items in een categorie."""
        category_key = f"category_{category_number}"
        if category_key in self.categories:
            return self.categories[category_key]["items"]
        return []
    
    def update_category_info(self, category_number, name=None, description=None, action=None, color=None):
        """Update categorie informatie."""
        category_key = f"category_{category_number}"
        if category_key in self.categories:
            if name:
                self.categories[category_key]["name"] = name
            if description:
                self.categories[category_key]["description"] = description
            if action:
                self.categories[category_key]["action"] = action
            if color:
                self.categories[category_key]["color"] = color
            self.save_categories()
            return True
        return False
    
    def export_to_config_format(self):
        """Exporteer naar config.py formaat."""
        result = {
            "CATEGORY_1_ITEMS": [],
            "CATEGORY_2_ITEMS": [],
            "CATEGORY_3_ITEMS": [],
            "CATEGORIES": {},
            "EMAIL_TEMPLATES": {}
        }
        
        for category_key, category_data in self.categories.items():
            category_num = int(category_key.split("_")[1])
            result[f"CATEGORY_{category_num}_ITEMS"] = category_data["items"]
            result["CATEGORIES"][category_num] = {
                "name": category_data["name"],
                "description": category_data["description"],
                "action": category_data["action"],
                "color": category_data["color"]
            }
        
        return result

def main():
    """Test de CategoryManager."""
    cm = CategoryManager()
    
    print("🚀 Category Manager Test")
    print("=" * 50)
    
    # Toon huidige categorieën
    for i in range(1, 4):
        print(f"\n📋 Categorie {i}: {cm.get_category_name(i)}")
        print(f"   Actie: {cm.get_category_action(i)}")
        print(f"   Items: {len(cm.get_all_items_in_category(i))}")
        print(f"   Voorbeelden: {cm.get_all_items_in_category(i)[:3]}")
    
    # Test item toevoegen
    print(f"\n➕ Voeg item 99999 toe aan categorie 1...")
    cm.add_item_to_category("99999", "category_1")
    
    # Test categorie bepaling
    print(f"\n🔍 Categorie voor item 10701: {cm.get_category_for_item('10701')}")
    print(f"🔍 Categorie voor item 10700: {cm.get_category_for_item('10700')}")
    print(f"🔍 Categorie voor item 99999: {cm.get_category_for_item('99999')}")
    
    print(f"\n✅ Category Manager test voltooid!")

if __name__ == "__main__":
    main() 