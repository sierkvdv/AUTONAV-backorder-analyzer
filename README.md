# Navision Backorder Analyzer

Een Python-script dat automatisch een Navision-backorder export analyseert en netjes visualiseert in een nieuwe Excel.

## 🚀 Snelle Start (Aanbevolen)

**Voor de eenvoudigste ervaring:**
1. **Dubbelklik** op `start_dashboard.bat`
2. **Sleep je Excel bestand** naar het dashboard of klik "📂 Bestand Kiezen"
3. **Klik "🔍 Analyse Starten"**
4. **Open het resultaat** met "📊 Output Openen"

Dat is het! 🎉

## 📊 Dashboard Features

### ✨ Gebruiksvriendelijke Interface
- **Moderne GUI** met emoji's en duidelijke knoppen
- **Real-time logging** - zie wat er gebeurt
- **Configuratie aanpassingen** direct in de interface
- **Automatische bestand opening** na analyse

### 🔧 Configuratie Opties
- **📍 Location Code**: Pas aan naar jouw locatie (standaard: DSV)
- **🔒 Fully Reserved**: Yes/No dropdown
- **📋 Order Status**: Pas aan naar jouw order status (standaard: Backorder)

### 📁 Bestand Beheer
- **📂 Bestand Kiezen**: Blader naar je Excel bestand
- **📊 Output Openen**: Open direct het resultaat
- **📄 Log Bestand Openen**: Bekijk gedetailleerde logs
- **📁 Output Map Openen**: Open de output map

## Functionaliteiten

### Fase 1: Basis Analyse en Visualisatie ✅
- **Data Import**: Leest Navision export (`.xlsx`) met pandas
- **Filtering**: Filtert op Location Code = DSV, Fully Reserved = No, Order Status = Backorder
- **Groepering**: Groepeert per Sales Order met splitsing in verzendbaar/backorder artikelen
- **Excel Output**: Creëert professionele Excel met kleurcodering en opmaak

### Fase 2: Actie-items en Workflow (Toekomstig)
- Automatische actie-items genereren
- Prioriteiten toekennen
- Workflow integratie

### Fase 3: E-mail Notificaties (Toekomstig)
- E-mail templates
- Automatische notificaties
- Dealer communicatie

## Vereiste Kolommen

De Navision export moet minimaal deze kolommen bevatten:
- `Sales Order No.` (ordernummer)
- `Item No.` (artikelnummer)
- `Description` (omschrijving)
- `Quantity` (besteld aantal)
- `Quantity Available` (beschikbaar aantal)
- `Location Code`
- `Fully Reserved` (Yes/No)
- `Order Status` (bijv. Backorder)
- `Customer Name` (dealer)

## Installatie

1. **Python installeren** (versie 3.8 of hoger)
2. **Dependencies installeren**:
   ```bash
   pip install -r requirements.txt
   ```

## Gebruik

### 🎯 Optie 1: Dashboard (Aanbevolen)
```bash
# Dubbelklik op start_dashboard.bat
# OF
python simple_dashboard.py
```

### 📝 Optie 2: Command Line
```bash
python backorder_analyzer.py
```

### 🖱️ Optie 3: Batch Files
- **Dashboard**: `start_dashboard.bat`
- **Command line**: `run_analyzer.bat`
- **PowerShell**: `run_analyzer.ps1`

## Output Format

De Excel output bevat:
- **Titel**: Met datum en tijdstip van analyse
- **Per Order**: Duidelijk gescheiden blokken
  - **Verzenden** (groene achtergrond): Artikelen met `Quantity Available > 0`
  - **Backorder** (rode achtergrond): Artikelen met `Quantity Available = 0`
- **Kolommen**: Item No., Description, Quantity, Available
- **Opmaak**: Professionele styling met borders, kleuren en uitlijning

## Configuratie

### Via Dashboard (Aanbevolen)
- Open het dashboard
- Pas de instellingen aan in de "⚙️ Configuratie" sectie
- Klik "🔍 Analyse Starten"

### Via Configuratie Bestand
Je kunt de volgende instellingen aanpassen in `config.py`:

```python
# Filter criteria
LOCATION_CODE = "DSV"           # Pas aan naar gewenste locatie
FULLY_RESERVED = "No"           # Filter op volledig gereserveerd
ORDER_STATUS = "Backorder"      # Filter op order status

# Kleuren (hex codes)
COLORS = {
    'header': '366092',         # Donkerblauw voor headers
    'sendable': 'C6EFCE',       # Lichtgroen voor verzendbare artikelen
    'backorder': 'FFC7CE',      # Lichtrood voor backorder artikelen
    'order_header': 'D9E1F2',   # Lichtblauw voor order headers
    'border': '000000'          # Zwart voor borders
}
```

## Troubleshooting

### Veelvoorkomende problemen:

1. **"Bestand niet gevonden"**
   - Controleer of het input bestand in de juiste map staat
   - Controleer de bestandsnaam in `INPUT_FILE`

2. **"Verplichte kolommen ontbreken"**
   - Controleer of alle vereiste kolommen aanwezig zijn
   - Let op hoofdlettergevoeligheid van kolomnamen

3. **"Geen data gevonden na filtering"**
   - Controleer of de filter criteria kloppen
   - Controleer de waarden in Location Code, Fully Reserved, Order Status

4. **Excel bestand kan niet geopend worden**
   - Zorg dat het bestand niet open staat in Excel
   - Controleer schrijfrechten in de output map

5. **Dashboard start niet**
   - Controleer of Python geïnstalleerd is
   - Controleer of alle dependencies geïnstalleerd zijn

## Logging

Het script genereert uitgebreide logs:
- **Dashboard**: Real-time logging in de interface
- **Log bestand**: `backorder_analyzer.log` met details
- **Console**: Directe output bij command line gebruik

## Uitbreidingen

Het script is modulair opgezet voor eenvoudige uitbreiding:

### Nieuwe filters toevoegen:
```python
def filter_backorder_data(df):
    # Voeg nieuwe filters toe
    filtered_df = filtered_df[filtered_df['Nieuwe_Kolom'] == 'Gewenste_Waarde']
    return filtered_df
```

### Nieuwe output formaten:
```python
def create_custom_output(orders):
    # Implementeer nieuwe output logica
    pass
```

## Bestanden Overzicht

### 🎯 Hoofdbestanden
- `simple_dashboard.py` - **Dashboard interface (aanbevolen)**
- `backorder_analyzer.py` - Hoofdscript voor analyse
- `config.py` - Configuratie instellingen

### 🚀 Launchers
- `start_dashboard.bat` - **Start dashboard (aanbevolen)**
- `run_analyzer.bat` - Start command line versie
- `run_analyzer.ps1` - PowerShell launcher

### 📊 Test & Documentatie
- `create_sample_data.py` - Genereert voorbeeld data
- `requirements.txt` - Python dependencies
- `README.md` - Deze documentatie

### 📁 Output
- `Output/Backorder_Analyse.xlsx` - **Gegenereerde analyse**
- `backorder_analyzer.log` - **Gedetailleerde logs**

## Technische Details

- **Python versie**: 3.8+
- **Hoofdbibliotheken**: pandas, openpyxl, tkinter
- **Bestandsformaten**: .xlsx (input en output)
- **Encoding**: UTF-8
- **Platform**: Windows, macOS, Linux

## Licentie

Dit script is ontwikkeld voor AUTONAV en is bedoeld voor intern gebruik.