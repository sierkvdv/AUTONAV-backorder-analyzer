#!/usr/bin/env python3
"""
Setup Script voor Gedeeld Gebruik
=================================

Dit script helpt bij het instellen van het backorder systeem voor gedeeld gebruik.
"""

import os
import shutil
import json
from pathlib import Path

def setup_shared_network():
    """Setup voor gedeeld gebruik via netwerkmap."""
    
    print("=== Setup voor Gedeeld Gebruik ===")
    print()
    
    # Vraag gebruiker om netwerkpad
    print("Voer het netwerkpad in waar het systeem gedeeld moet worden:")
    print("Voorbeeld: \\\\server\\shared\\AUTONAV of Z:\\AUTONAV")
    network_path = input("Netwerkpad: ").strip()
    
    if not network_path:
        print("❌ Geen netwerkpad ingevoerd. Setup geannuleerd.")
        return
    
    # Maak netwerkmap aan als deze niet bestaat
    try:
        Path(network_path).mkdir(parents=True, exist_ok=True)
        print(f"✅ Netwerkmap aangemaakt/gecontroleerd: {network_path}")
    except Exception as e:
        print(f"❌ Fout bij aanmaken netwerkmap: {e}")
        return
    
    # Kopieer bestanden naar netwerkmap
    files_to_copy = [
        'backorder_analyzer.py',
        'simple_dashboard.py', 
        'category_manager_gui.py',
        'category_manager.py',
        'config.py',
        'requirements.txt',
        'README.md',
        'start_dashboard.bat',
        'run_analyzer.bat',
        'run_analyzer.ps1'
    ]
    
    print("\n📁 Bestanden kopiëren naar netwerkmap...")
    for file in files_to_copy:
        if os.path.exists(file):
            try:
                shutil.copy2(file, network_path)
                print(f"  ✅ {file}")
            except Exception as e:
                print(f"  ❌ {file}: {e}")
        else:
            print(f"  ⚠️  {file} (niet gevonden)")
    
    # Maak Output map aan
    output_path = os.path.join(network_path, 'Output')
    try:
        Path(output_path).mkdir(exist_ok=True)
        print(f"✅ Output map aangemaakt: {output_path}")
    except Exception as e:
        print(f"❌ Fout bij aanmaken Output map: {e}")
    
    # Kopieer configuratie bestanden als ze bestaan
    config_files = ['category_config.json', 'item_links.json']
    for config_file in config_files:
        if os.path.exists(config_file):
            try:
                shutil.copy2(config_file, network_path)
                print(f"✅ Configuratie gekopieerd: {config_file}")
            except Exception as e:
                print(f"❌ Fout bij kopiëren {config_file}: {e}")
    
    # Maak batch bestand voor eenvoudige toegang
    create_network_batch_file(network_path)
    
    print("\n" + "="*50)
    print("✅ Setup voltooid!")
    print(f"📁 Systeem beschikbaar op: {network_path}")
    print("\n📋 Volgende stappen:")
    print("1. Deel het netwerkpad met collega's")
    print("2. Collega's kunnen 'start_dashboard.bat' uitvoeren")
    print("3. Configuratie wordt automatisch gedeeld")
    print("\n💡 Tips:")
    print("- Zorg dat alle gebruikers schrijfrechten hebben op de netwerkmap")
    print("- Maak regelmatig backups van category_config.json en item_links.json")
    print("- Test het systeem eerst met een kleine groep gebruikers")

def create_network_batch_file(network_path):
    """Maak een batch bestand voor eenvoudige toegang."""
    
    batch_content = f"""@echo off
echo ========================================
echo    AUTONAV Backorder Analyzer
echo ========================================
echo.
echo Start dashboard...
echo.

REM Controleer of Python geïnstalleerd is
python --version >nul 2>&1
if errorlevel 1 (
    echo ❌ Python is niet geïnstalleerd!
    echo Download Python van: https://www.python.org/downloads/
    pause
    exit /b 1
)

REM Installeer dependencies als nodig
if not exist "requirements_installed.txt" (
    echo 📦 Installeer dependencies...
    python -m pip install -r requirements.txt
    echo. > requirements_installed.txt
)

REM Start dashboard
echo 🚀 Start dashboard...
python simple_dashboard.py

pause
"""
    
    batch_file = os.path.join(network_path, 'start_dashboard.bat')
    try:
        with open(batch_file, 'w', encoding='utf-8') as f:
            f.write(batch_content)
        print(f"✅ Batch bestand aangemaakt: {batch_file}")
    except Exception as e:
        print(f"❌ Fout bij aanmaken batch bestand: {e}")

def create_user_guide():
    """Maak een gebruikershandleiding."""
    
    guide_content = """# Gebruikershandleiding - AUTONAV Backorder Analyzer

## 🚀 Snel Starten

1. **Open de netwerkmap** waar het systeem is geïnstalleerd
2. **Dubbelklik** op `start_dashboard.bat`
3. **Wacht** tot het dashboard opent
4. **Upload** je Navision export Excel bestand
5. **Klik** op "🔍 Analyse Starten"

## 📋 Categorieën Beheren

### Artikelen Toevoegen aan Categorieën:
1. Klik op "📋 Categorie Manager" in het dashboard
2. Selecteer een categorie (1, 2, of 3)
3. Voer het artikelnummer in
4. Klik "➕ Artikel Toevoegen"

### Links Toevoegen per Artikel:
1. In Categorie Manager, ga naar "🔗 Links per Artikel"
2. Selecteer het artikelnummer
3. Voer de fabrikant link in
4. Voer de externe verkoper link in
5. Klik "💾 Links Opslaan"

## 📊 Resultaten Bekijken

Na de analyse vind je:
- **Backorder_Analyse_vX.xlsx**: Hoofdanalyse met gesplitste orders
- **Backorder_Analyse_vX_Emails.xlsx**: E-mail content voor verzending

## 🔧 Troubleshooting

### Python niet gevonden:
- Download Python van: https://www.python.org/downloads/
- Zorg dat "Add Python to PATH" is aangevinkt

### Geen schrijfrechten:
- Vraag de IT-beheerder om schrijfrechten op de netwerkmap

### Dashboard opent niet:
- Controleer of alle bestanden aanwezig zijn
- Probeer het batch bestand als administrator uit te voeren

## 📞 Support

Voor vragen of problemen, neem contact op met de systeembeheerder.
"""
    
    try:
        with open('Gebruikershandleiding.md', 'w', encoding='utf-8') as f:
            f.write(guide_content)
        print("✅ Gebruikershandleiding aangemaakt: Gebruikershandleiding.md")
    except Exception as e:
        print(f"❌ Fout bij aanmaken handleiding: {e}")

if __name__ == "__main__":
    print("AUTONAV Backorder Analyzer - Setup voor Gedeeld Gebruik")
    print("=" * 60)
    print()
    
    choice = input("Wat wilt u doen?\n1. Setup voor netwerkmap\n2. Maak gebruikershandleiding\n3. Beide\nKeuze (1-3): ").strip()
    
    if choice in ['1', '3']:
        setup_shared_network()
    
    if choice in ['2', '3']:
        create_user_guide()
    
    print("\nSetup voltooid! 🎉") 