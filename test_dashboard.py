#!/usr/bin/env python3
"""
Test Dashboard
=============

Eenvoudige test om te controleren of tkinter werkt.
"""

import tkinter as tk
from tkinter import ttk

def main():
    """Test tkinter."""
    print("🚀 Tkinter test gestart...")
    
    try:
        # Maak root window
        root = tk.Tk()
        root.title("Tkinter Test")
        root.geometry("400x300")
        
        # Voeg een label toe
        label = ttk.Label(root, text="Tkinter werkt! 🎉", font=("Arial", 16))
        label.pack(pady=50)
        
        # Voeg een knop toe
        button = ttk.Button(root, text="Sluiten", command=root.quit)
        button.pack(pady=20)
        
        print("✅ Tkinter window geopend")
        
        # Start main loop
        root.mainloop()
        
        print("✅ Tkinter test voltooid")
        
    except Exception as e:
        print(f"❌ Tkinter fout: {e}")

if __name__ == "__main__":
    main() 