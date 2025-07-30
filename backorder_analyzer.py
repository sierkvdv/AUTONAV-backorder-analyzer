#!/usr/bin/env python3
"""
Navision Backorder Analyzer
===========================

Dit script analyseert Navision backorder exports en genereert een overzichtelijke Excel
met verzendbare en backorder artikelen, inclusief automatische categorisering en e-mailgeneratie.

Auteur: AUTONAV
Versie: 2.0
Datum: 2024
"""

import pandas as pd
import numpy as np
import logging
import os
from datetime import datetime
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils.dataframe import dataframe_to_rows

# Import configuratie
try:
    from config import *
except ImportError:
    # Fallback configuratie als config.py niet bestaat
    INPUT_FILE = "Navision_Backorder_Export_Realistic.xlsx"
    OUTPUT_FILE = "Output/Backorder_Analyse.xlsx"
    LOCATION_CODE = "DSV"
    FULLY_RESERVED = "No"
    ORDER_STATUS = "Backorder"
    COLORS = {
        'header': '366092',
        'sendable': 'C6EFCE',
        'backorder': 'FFC7CE',
        'order_header': 'D9E1F2',
        'border': '000000',
        'category_1': 'FF6B6B',
        'category_2': '4ECDC4',
        'category_3': 'FFA500'
    }
    COLUMN_WIDTHS = [15, 40, 12, 12, 15, 40, 12, 12, 20]
    LOG_FILE = "backorder_analyzer.log"
    LOG_LEVEL = "INFO"
    REQUIRED_COLUMNS = [
        'Sales Order No.', 'Item No.', 'Description', 'Quantity',
        'Quantity Available', 'Location Code', 'Fully Reserved',
        'Order Status', 'Customer Name'
    ]
    # Categorie configuratie
    CATEGORY_1_ITEMS = ['OIL-5W30-1L', 'OIL-10W40-1L', 'BRAKE-FLUID', 'COOLANT-1L']
    CATEGORY_2_ITEMS = ['BRAKE-FLUID-DOT4', 'BRAKE-PADS-FRONT', 'BRAKE-PADS-REAR']
    CATEGORY_3_ITEMS = ['EXHAUST-CAT', 'EXHAUST-MUFFLER', 'EXHAUST-PIPE']
    CATEGORIES = {
        1: {'name': 'Bestel bij fabrikant', 'color': 'FF6B6B'},
        2: {'name': 'Binnenkort leverbaar', 'color': '4ECDC4'},
        3: {'name': 'Geen voorraadvooruitzicht', 'color': 'FFA500'}
    }
    EMAIL_TEMPLATES = {}
    SALESFORCE_EMAIL_SETTINGS = {'enabled': False}

# Import CategoryManager
try:
    from category_manager import CategoryManager
    category_manager = CategoryManager()
except ImportError:
    category_manager = None

# Setup logging
logging.basicConfig(
    level=getattr(logging, LOG_LEVEL),
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler(LOG_FILE, encoding='utf-8'),
        logging.StreamHandler()
    ]
)

def load_navision_data(file_path):
    """Laad de Navision export data."""
    logging.info(f"Laden van Navision export: {file_path}")
    
    try:
        df = pd.read_excel(file_path)
        logging.info(f"Data geladen: {len(df)} rijen, {len(df.columns)} kolommen")
        return df
    except Exception as e:
        logging.error(f"Fout bij laden van bestand: {e}")
        raise

def validate_columns(df):
    """Valideer of alle verplichte kolommen aanwezig zijn."""
    missing_columns = [col for col in REQUIRED_COLUMNS if col not in df.columns]
    if missing_columns:
        raise ValueError(f"Ontbrekende kolommen: {missing_columns}")
    logging.info("Alle verplichte kolommen aanwezig")

def filter_backorder_data(df):
    """Filter de data op basis van de criteria."""
    original_count = len(df)
    
    # Filter op Location Code
    df = df[df['Location Code'] == LOCATION_CODE]
    logging.info(f"Na Location Code filter ({LOCATION_CODE}): {len(df)} rijen")
    
    # Filter op Fully Reserved
    df = df[df['Fully Reserved'] == FULLY_RESERVED]
    logging.info(f"Na Fully Reserved filter ({FULLY_RESERVED}): {len(df)} rijen")
    
    # Filter op Order Status
    df = df[df['Order Status'] == ORDER_STATUS]
    logging.info(f"Na Order Status filter ({ORDER_STATUS}): {len(df)} rijen")
    
    logging.info(f"Filtering voltooid: {original_count} -> {len(df)} rijen")
    return df

def categorize_backorder_items(df):
    """Categoriseer backorder artikelen in de drie categorieën."""
    def get_category(item_no):
        if category_manager:
            return category_manager.get_category_for_item(item_no)
        else:
            # Fallback naar config.py categorieën
            if item_no in CATEGORY_1_ITEMS:
                return 1
            elif item_no in CATEGORY_2_ITEMS:
                return 2
            elif item_no in CATEGORY_3_ITEMS:
                return 3
            else:
                return 2  # Default naar categorie 2
    
    # Voeg categorie toe aan DataFrame
    df['Category'] = df['Item No.'].apply(get_category)
    
    # Bepaal categorie namen
    if category_manager:
        df['Category_Name'] = df['Category'].apply(lambda x: category_manager.get_category_name(x))
        df['Category_Action'] = df['Category'].apply(lambda x: category_manager.get_category_action(x))
    else:
        df['Category_Name'] = df['Category'].map({1: CATEGORIES[1]['name'], 
                                                 2: CATEGORIES[2]['name'], 
                                                 3: CATEGORIES[3]['name']})
        df['Category_Action'] = df['Category'].map({1: CATEGORIES[1].get('action', 'Onbekend'), 
                                                   2: CATEGORIES[2].get('action', 'Onbekend'), 
                                                   3: CATEGORIES[3].get('action', 'Onbekend')})
    
    return df

def group_by_sales_order(df):
    """Groepeer data per Sales Order en categoriseer artikelen."""
    grouped_data = {}
    
    for order_no in df['Sales Order No.'].unique():
        order_data = df[df['Sales Order No.'] == order_no]
        customer_name = order_data['Customer Name'].iloc[0]
        
        # Split in verzendbaar en backorder
        sendable = order_data[order_data['Quantity Available'] > 0]
        backorder = order_data[order_data['Quantity Available'] == 0]
        
        # Categoriseer backorder artikelen
        if len(backorder) > 0:
            backorder = categorize_backorder_items(backorder)
        
        grouped_data[order_no] = {
            'customer': customer_name,
            'sendable': sendable,
            'backorder': backorder,
            'total_items': len(order_data),
            'sendable_count': len(sendable),
            'backorder_count': len(backorder)
        }
        
        logging.info(f"Order {order_no} ({customer_name}): {len(sendable)} verzendbaar, {len(backorder)} backorder")
    
    return grouped_data

def generate_email_content(item_data, category):
    """Genereer e-mail content voor een artikel."""
    if category not in EMAIL_TEMPLATES:
        return None
    
    template = EMAIL_TEMPLATES[category]
    
    # Vul template variabelen
    email_data = {
        'customer_name': item_data['Customer Name'],
        'item_no': item_data['Item No.'],
        'item_description': item_data['Description'],
        'order_no': item_data['Sales Order No.'],
        'quantity': item_data['Quantity']
    }
    
    # Voeg specifieke links toe via CategoryManager
    if category_manager:
        if category == 1:
            # Categorie 1: Link naar fabrikant
            manufacturer_link = category_manager.get_item_link(item_data['Item No.'], 'fabrikant')
            if manufacturer_link:
                email_data['manufacturer_link'] = manufacturer_link
            else:
                email_data['manufacturer_link'] = 'https://www.original-equipment-parts.com'
        elif category == 3:
            # Categorie 3: Link naar externe verkoper
            external_link = category_manager.get_item_link(item_data['Item No.'], 'externe_verkoper')
            if external_link:
                email_data['external_seller_link'] = external_link
            else:
                email_data['external_seller_link'] = 'https://www.autodoc.nl'
    else:
        # Fallback naar oude methode
        if category == 1:
            manufacturer_links = template.get('manufacturer_links', {})
            email_data['manufacturer_link'] = manufacturer_links.get(
                item_data['Item No.'], 
                manufacturer_links.get('default', 'https://www.original-equipment-parts.com')
            )
        elif category == 3:
            external_links = template.get('external_seller_links', {})
            email_data['external_seller_link'] = external_links.get(
                item_data['Item No.'], 
                external_links.get('default', 'https://www.autodoc.nl')
            )
    
    # Genereer e-mail content
    subject = template['subject'].format(**email_data)
    body = template['body'].format(**email_data)
    
    return {
        'to': item_data['Customer Name'],
        'subject': subject,
        'body': body.strip(),
        'category': category,
        'item_data': item_data
    }

def create_excel_workbook(grouped_data):
    """Maak een Excel werkboek met de geanalyseerde data."""
    logging.info("Excel werkboek succesvol aangemaakt")
    
    wb = Workbook()
    ws = wb.active
    ws.title = "Backorder Analyse"
    
    # Definieer kleuren
    colors = {
        'header': PatternFill(start_color=COLORS['header'], end_color=COLORS['header'], fill_type='solid'),
        'sendable': PatternFill(start_color=COLORS['sendable'], end_color=COLORS['sendable'], fill_type='solid'),
        'backorder': PatternFill(start_color=COLORS['backorder'], end_color=COLORS['backorder'], fill_type='solid'),
        'order_header': PatternFill(start_color=COLORS['order_header'], end_color=COLORS['order_header'], fill_type='solid'),
        'category_1': PatternFill(start_color=COLORS['category_1'], end_color=COLORS['category_1'], fill_type='solid'),
        'category_2': PatternFill(start_color=COLORS['category_2'], end_color=COLORS['category_2'], fill_type='solid'),
        'category_3': PatternFill(start_color=COLORS['category_3'], end_color=COLORS['category_3'], fill_type='solid')
    }
    
    # Definieer borders
    thin_border = Border(
        left=Side(style='thin'),
        right=Side(style='thin'),
        top=Side(style='thin'),
        bottom=Side(style='thin')
    )
    
    # Definieer fonts
    header_font = Font(bold=True, color="FFFFFF")
    normal_font = Font()
    
    current_row = 1
    
    # Voor elke order
    for order_no, order_info in grouped_data.items():
        customer = order_info['customer']
        
        # Order header
        ws.cell(row=current_row, column=1, value=f"Order: {order_no}")
        ws.cell(row=current_row, column=2, value=f"Klant: {customer}")
        ws.cell(row=current_row, column=3, value=f"Totaal: {order_info['total_items']}")
        ws.cell(row=current_row, column=4, value=f"Verzendbaar: {order_info['sendable_count']}")
        ws.cell(row=current_row, column=5, value=f"Backorder: {order_info['backorder_count']}")
        
        # Styling voor order header
        for col in range(1, 6):
            cell = ws.cell(row=current_row, column=col)
            cell.fill = colors['order_header']
            cell.font = Font(bold=True)
            cell.border = thin_border
        
        current_row += 1
        
        # Verzendbare artikelen
        if len(order_info['sendable']) > 0:
            ws.cell(row=current_row, column=1, value="✅ VERZENDBAAR")
            ws.cell(row=current_row, column=1).fill = colors['sendable']
            ws.cell(row=current_row, column=1).font = Font(bold=True)
            current_row += 1
            
            # Headers voor verzendbare artikelen
            headers = ['Artikelnummer', 'Omschrijving', 'Aantal', 'Beschikbaar']
            for col, header in enumerate(headers, 1):
                cell = ws.cell(row=current_row, column=col, value=header)
                cell.fill = colors['header']
                cell.font = header_font
                cell.border = thin_border
            current_row += 1
            
            # Verzendbare artikelen data
            for _, item in order_info['sendable'].iterrows():
                ws.cell(row=current_row, column=1, value=item['Item No.'])
                ws.cell(row=current_row, column=2, value=item['Description'])
                ws.cell(row=current_row, column=3, value=item['Quantity'])
                ws.cell(row=current_row, column=4, value=item['Quantity Available'])
                
                # Styling
                for col in range(1, 5):
                    cell = ws.cell(row=current_row, column=col)
                    cell.fill = colors['sendable']
                    cell.border = thin_border
                current_row += 1
        
        # Backorder artikelen
        if len(order_info['backorder']) > 0:
            ws.cell(row=current_row, column=1, value="❌ BACKORDER")
            ws.cell(row=current_row, column=1).fill = colors['backorder']
            ws.cell(row=current_row, column=1).font = Font(bold=True)
            current_row += 1
            
            # Headers voor backorder artikelen
            headers = ['Artikelnummer', 'Omschrijving', 'Aantal', 'Beschikbaar', 'Categorie', 'Actie']
            for col, header in enumerate(headers, 1):
                cell = ws.cell(row=current_row, column=col, value=header)
                cell.fill = colors['header']
                cell.font = header_font
                cell.border = thin_border
            current_row += 1
            
            # Backorder artikelen data
            for _, item in order_info['backorder'].iterrows():
                category = item['Category']
                category_name = item['Category_Name']
                
                ws.cell(row=current_row, column=1, value=item['Item No.'])
                ws.cell(row=current_row, column=2, value=item['Description'])
                ws.cell(row=current_row, column=3, value=item['Quantity'])
                ws.cell(row=current_row, column=4, value=item['Quantity Available'])
                ws.cell(row=current_row, column=5, value=category_name)
                
                # Actie kolom
                if category == 1:
                    action = "Verwijder + E-mail naar fabrikant"
                elif category == 2:
                    action = "Behoud backorder"
                elif category == 3:
                    action = "Verwijder + E-mail naar externe verkoper"
                else:
                    action = "Onbekend"
                
                ws.cell(row=current_row, column=6, value=action)
                
                # Styling met categorie kleur
                category_color = colors[f'category_{category}']
                for col in range(1, 7):
                    cell = ws.cell(row=current_row, column=col)
                    cell.fill = category_color
                    cell.border = thin_border
                current_row += 1
        
        # Spacing tussen orders
        current_row += 2
    
    # Auto-fit kolommen (veilige versie)
    for col in range(1, 7):  # Max 6 kolommen
        max_length = 0
        column_letter = ws.cell(row=1, column=col).column_letter
        
        for row in range(1, current_row):
            cell = ws.cell(row=row, column=col)
            if cell.value:
                try:
                    if len(str(cell.value)) > max_length:
                        max_length = len(str(cell.value))
                except:
                    pass
        
        adjusted_width = min(max_length + 2, 50)
        ws.column_dimensions[column_letter].width = adjusted_width
    
    return wb

def generate_email_report(grouped_data):
    """Genereer een rapport van alle e-mails die verzonden moeten worden."""
    email_report = []
    
    for order_no, order_info in grouped_data.items():
        for _, item in order_info['backorder'].iterrows():
            category = item['Category']
            
            # Alleen categorie 1 en 3 krijgen e-mails
            if category in [1, 3]:
                email_content = generate_email_content(item, category)
                if email_content:
                    email_report.append(email_content)
    
    return email_report

def save_excel_file(wb, file_path):
    """Sla het Excel bestand op."""
    # Maak output directory als deze niet bestaat
    os.makedirs(os.path.dirname(file_path), exist_ok=True)
    
    wb.save(file_path)
    logging.info(f"Excel bestand opgeslagen: {file_path}")

def save_email_report(email_report, file_path):
    """Sla het e-mail rapport op als Excel bestand."""
    if not email_report:
        logging.info("Geen e-mails om te verzenden")
        return
    
    # Maak DataFrame van e-mails
    email_data = []
    for email in email_report:
        email_data.append({
            'Ordernummer': email['item_data']['Sales Order No.'],
            'Klant': email['to'],
            'Artikelnummer': email['item_data']['Item No.'],
            'Omschrijving': email['item_data']['Description'],
            'Aantal': email['item_data']['Quantity'],
            'Categorie': email['category'],
            'Onderwerp': email['subject'],
            'E-mail Body': email['body']
        })
    
    df = pd.DataFrame(email_data)
    
    # Sla op als Excel
    with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
        df.to_excel(writer, sheet_name='E-mails', index=False)
        
        # Auto-fit kolommen
        worksheet = writer.sheets['E-mails']
        for column in worksheet.columns:
            max_length = 0
            column_letter = column[0].column_letter
            
            for cell in column:
                try:
                    if len(str(cell.value)) > max_length:
                        max_length = len(str(cell.value))
                except:
                    pass
            
            adjusted_width = min(max_length + 2, 50)
            worksheet.column_dimensions[column_letter].width = adjusted_width
    
    logging.info(f"E-mail rapport opgeslagen: {file_path}")

def main():
    """Hoofdfunctie van het script."""
    logging.info("=== Navision Backorder Analyzer gestart ===")
    
    try:
        # Laad data
        df = load_navision_data(INPUT_FILE)
        
        # Valideer kolommen
        validate_columns(df)
        
        # Filter data
        filtered_df = filter_backorder_data(df)
        
        # Groepeer per order
        grouped_data = group_by_sales_order(filtered_df)
        
        # Maak Excel werkboek
        wb = create_excel_workbook(grouped_data)
        
        # Sla Excel op
        save_excel_file(wb, OUTPUT_FILE)
        
        # Genereer e-mail rapport
        email_report = generate_email_report(grouped_data)
        if email_report:
            email_file = OUTPUT_FILE.replace('.xlsx', '_Emails.xlsx')
            save_email_report(email_report, email_file)
        
        # Print samenvatting
        total_orders = len(grouped_data)
        total_sendable = sum(order['sendable_count'] for order in grouped_data.values())
        total_backorder = sum(order['backorder_count'] for order in grouped_data.values())
        
        logging.info("=== Analyse voltooid ===")
        logging.info(f"Totaal orders: {total_orders}")
        logging.info(f"Totaal verzendbare artikelen: {total_sendable}")
        logging.info(f"Totaal backorder artikelen: {total_backorder}")
        logging.info(f"Output bestand: {OUTPUT_FILE}")
        
        if email_report:
            logging.info(f"E-mails om te verzenden: {len(email_report)}")
            logging.info(f"E-mail rapport: {email_file}")
        
        # Categorie statistieken
        category_counts = {1: 0, 2: 0, 3: 0}
        for order_info in grouped_data.values():
            for _, item in order_info['backorder'].iterrows():
                category_counts[item['Category']] += 1
        
        logging.info("Backorder categorieën:")
        for cat_num, count in category_counts.items():
            if count > 0:
                cat_name = CATEGORIES[cat_num]['name']
                logging.info(f"  - {cat_name}: {count} artikelen")
        
    except Exception as e:
        logging.error(f"Fout tijdens uitvoering: {e}")
        raise

if __name__ == "__main__":
    main()