#!/usr/bin/env python3
"""
Script om test artikelen toe te voegen aan categorieën.
"""

from category_manager import CategoryManager

def add_test_items():
    """Voeg test artikelen toe aan categorieën."""
    cm = CategoryManager()
    
    print("🔧 Toevoegen van test artikelen aan categorieën")
    print("=" * 60)
    
    # Test artikelen uit Navision_Backorder_Export_QWIC_v2.xlsx
    test_items = [
        10700, 10701, 10702, 10703, 10704, 10706, 10708, 10709, 10710,
        10713, 10714, 10716, 10717, 10718, 10720
    ]
    
    # Voorbeeld categorisering - pas aan op basis van echte business rules
    # Categorie 1: Bestel bij fabrikant (artikelen die niet meer leverbaar zijn via QWIC)
    category_1_items = [
        10700, 10701, 10702  # Voorbeelden - pas aan op basis van echte data
    ]
    
    # Categorie 2: Binnenkort leverbaar (artikelen die QWIC nog wel kan leveren)
    category_2_items = [
        10703, 10704, 10706  # Voorbeelden - pas aan op basis van echte data
    ]
    
    # Categorie 3: Geen voorraadvooruitzicht (artikelen die niet meer beschikbaar zijn)
    category_3_items = [
        10708, 10709, 10710  # Voorbeelden - pas aan op basis van echte data
    ]
    
    # Categorie 4: Vervang door alternatief (artikelen die vervangen worden door QWIC alternatief)
    category_4_items = [
        10713, 10714, 10716, 10717, 10718, 10720  # Voorbeelden - pas aan op basis van echte data
    ]
    
    # Voeg artikelen toe aan categorie 1
    print("📦 Toevoegen aan Categorie 1 (Bestel bij fabrikant):")
    for item in category_1_items:
        if cm.add_item_to_category(str(item), "category_1"):
            print(f"  ✅ {item} toegevoegd")
        else:
            print(f"  ⚠️  {item} was al toegevoegd")
    
    # Voeg artikelen toe aan categorie 2
    print("\n📦 Toevoegen aan Categorie 2 (Binnenkort leverbaar):")
    for item in category_2_items:
        if cm.add_item_to_category(str(item), "category_2"):
            print(f"  ✅ {item} toegevoegd")
        else:
            print(f"  ⚠️  {item} was al toegevoegd")
    
    # Voeg artikelen toe aan categorie 3
    print("\n📦 Toevoegen aan Categorie 3 (Geen voorraadvooruitzicht):")
    for item in category_3_items:
        if cm.add_item_to_category(str(item), "category_3"):
            print(f"  ✅ {item} toegevoegd")
        else:
            print(f"  ⚠️  {item} was al toegevoegd")
    
    # Voeg artikelen toe aan categorie 4
    print("\n📦 Toevoegen aan Categorie 4 (Vervang door alternatief):")
    for item in category_4_items:
        if cm.add_item_to_category(str(item), "category_4"):
            print(f"  ✅ {item} toegevoegd")
        else:
            print(f"  ⚠️  {item} was al toegevoegd")
    
    print("\n📋 Nieuwe categorie configuratie:")
    print("=" * 60)
    
    for i in range(1, 5):
        items = cm.get_all_items_in_category(i)
        name = cm.get_category_name(i)
        print(f"Categorie {i} ({name}): {items}")
    
    print("\n✅ Klaar! De test artikelen zijn toegevoegd aan de categorieën.")
    print("💡 Pas de artikelnummers aan op basis van echte business rules.")

if __name__ == "__main__":
    add_test_items() 