#!/usr/bin/env python3
"""
Test Category Manager GUI
"""

from category_manager import CategoryManager

def test_category_manager():
    print("=== TEST CATEGORY MANAGER ===")
    
    cm = CategoryManager()
    
    print(f"Categories loaded: {list(cm.categories.keys())}")
    
    for i in range(1, 4):
        category_key = f"category_{i}"
        print(f"\n--- Categorie {i} ---")
        print(f"Key: {category_key}")
        print(f"Name: {cm.get_category_name(i)}")
        print(f"Action: {cm.get_category_action(i)}")
        print(f"Color: {cm.get_category_color(i)}")
        print(f"Items: {cm.get_all_items_in_category(i)}")
        
        # Check direct access
        if category_key in cm.categories:
            data = cm.categories[category_key]
            print(f"Direct data - Name: {data.get('name', 'MISSING')}")
            print(f"Direct data - Action: {data.get('action', 'MISSING')}")
            print(f"Direct data - Color: {data.get('color', 'MISSING')}")
        else:
            print("❌ Category key not found!")

if __name__ == "__main__":
    test_category_manager() 