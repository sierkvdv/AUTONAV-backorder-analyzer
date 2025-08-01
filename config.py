# Navision Backorder Analyzer - Configuratie
# =========================================

# =============================================================================
# BESTANDSINSTELLINGEN
# =============================================================================

# Bestandsnaam van de Navision export (pas dit aan naar jouw bestand)
INPUT_FILE = "Navision_Backorder_Export_QWIC_v2.xlsx"

# Output bestandsnaam
OUTPUT_FILE = "Output/Backorder_Analyse_v7.xlsx"

# =============================================================================
# FILTER CRITERIA
# =============================================================================

# Location Code filter (leeg = alle locaties)
LOCATION_CODE = ""

# Fully Reserved filter (leeg = alle items)
FULLY_RESERVED = ""

# Order Status filter (leeg = alle statussen)
ORDER_STATUS = ""

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
CATEGORIES = {
    1: {
        'name': 'Bestel bij fabrikant',
        'description': 'Artikel is niet meer via QWIC leverbaar. Backorder wordt verwijderd en dealer krijgt automatische e-mail via Salesforce met link naar fabrikant.',
        'action': 'Verwijder backorder + E-mail naar fabrikant',
        'color': 'FF6B6B'  # Rood
    },
    2: {
        'name': 'Binnenkort leverbaar',
        'description': 'Backorder blijft staan en wordt niet gemaild (Navision stuurt automatisch zodra artikel binnen is).',
        'action': 'Behoud backorder',
        'color': '4ECDC4'  # Turquoise
    },
    3: {
        'name': 'Geen voorraadvooruitzicht',
        'description': 'Artikel komt niet meer of pas over zeer lange tijd. Backorder wordt verwijderd en dealer krijgt e-mail met link naar externe verkoper.',
        'action': 'Verwijder backorder + E-mail naar externe verkoper',
        'color': 'FFA500'  # Oranje
    }
}

# =============================================================================
# E-MAIL INSTELLINGEN
# =============================================================================

# E-mail templates voor elke categorie
EMAIL_TEMPLATES = {
    1: {
        'subject': 'Backorder artikel niet meer leverbaar via QWIC - {item_description}',
        'body': """
Beste {customer_name},

Uw backorder artikel "{item_description}" (Artikelnummer: {item_no}) is helaas niet meer leverbaar via QWIC.

Wij raden u aan om dit artikel direct bij de fabrikant te bestellen:
🔗 {manufacturer_link}

Ordergegevens:
- Ordernummer: {order_no}
- Artikel: {item_description}
- Aantal: {quantity}

Voor vragen kunt u contact opnemen met onze klantenservice.

Met vriendelijke groet,
QWIC Team
        """,
        'manufacturer_links': {
            'OIL-5W30-1L': 'https://www.castrol.com/nl_nl/netherlands/products/castrol-edge-5w-30.html',
            'OIL-10W40-1L': 'https://www.mobil.com/nl-nl/passenger-vehicle-lube/pds/gl-nl-mobil-super-2000-10w-40',
            'BRAKE-FLUID': 'https://www.bosch-automotive-catalog.com/nl/nl/brake-fluid-dot4',
            'AIR-FILTER': 'https://www.mann-filter.com/nl/nl/air-filters',
            'OIL-FILTER': 'https://www.mann-filter.com/nl/nl/oil-filters',
            'default': 'https://www.original-equipment-parts.com'
        }
    },
    3: {
        'subject': 'Backorder artikel niet meer beschikbaar - {item_description}',
        'body': """
Beste {customer_name},

Uw backorder artikel "{item_description}" (Artikelnummer: {item_no}) is helaas niet meer beschikbaar via QWIC.

Wij raden u aan om dit artikel bij een externe verkoper te bestellen:
🔗 {external_seller_link}

Ordergegevens:
- Ordernummer: {order_no}
- Artikel: {item_description}
- Aantal: {quantity}

Voor vragen kunt u contact opnemen met onze klantenservice.

Met vriendelijke groet,
QWIC Team
        """,
        'external_seller_links': {
            'EXHAUST-CAT': 'https://www.autodoc.nl/uitlaat/katalysator',
            'EXHAUST-MUFFLER': 'https://www.autodoc.nl/uitlaat/uitlaatdemper',
            'STEERING-WHEEL': 'https://www.autodoc.nl/stuur/stuurwiel',
            'default': 'https://www.autodoc.nl'
        }
    }
}

# Salesforce e-mail instellingen
SALESFORCE_EMAIL_SETTINGS = {
    'enabled': True,
    'from_email': 'backorder@qwic.nl',
    'from_name': 'QWIC Backorder Systeem',
    'reply_to': 'klantenservice@qwic.nl',
    'cc': ['backorder-admin@qwic.nl'],
    'bcc': []
}

# =============================================================================
# EXCEL OPMAAK KLEUREN
# =============================================================================

# Kleuren voor de Excel output (hex codes)
COLORS = {
    'header': '366092',      # Donkerblauw voor headers
    'sendable': 'C6EFCE',    # Lichtgroen voor verzendbare artikelen
    'backorder': 'FFC7CE',   # Lichtrood voor backorder artikelen
    'order_header': 'D9E1F2', # Lichtblauw voor order headers
    'border': '000000',      # Zwart voor borders
    'category_1': 'FF6B6B',  # Rood voor categorie 1
    'category_2': '4ECDC4',  # Turquoise voor categorie 2
    'category_3': 'FFA500'   # Oranje voor categorie 3
}

# =============================================================================
# KOLOM BREEDTES
# =============================================================================

# Kolom breedtes voor Excel output
COLUMN_WIDTHS = [15, 40, 12, 12, 15, 40, 12, 12, 20]

# =============================================================================
# LOGGING INSTELLINGEN
# =============================================================================

# Log bestandsnaam
LOG_FILE = "backorder_analyzer.log"

# Log niveau (DEBUG, INFO, WARNING, ERROR, CRITICAL)
LOG_LEVEL = "INFO"

# =============================================================================
# VERPLICHTE KOLOMMEN
# =============================================================================

# Lijst van verplichte kolommen in de Navision export
REQUIRED_COLUMNS = [
    'Sales Order No.',
    'Item No.',
    'Description',
    'Quantity',
    'Quantity Available',
    'Location Code',
    'Fully Reserved',
    'Order Status',
    'Customer Name'
]

# =============================================================================
# FASE 2 & 3 INSTELLINGEN (toekomstig)
# =============================================================================

# Actie-item instellingen voor Fase 2
ACTION_SETTINGS = {
    'priority_levels': ['Laag', 'Normaal', 'Hoog', 'Kritiek'],
    'auto_generate_actions': True,
    'include_customer_contact': True
}