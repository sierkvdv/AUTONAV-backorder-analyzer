# 🚗 AUTONAV Backorder Analyzer - Gebruikershandleiding

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

---

## 🛠️ Voor Systeembeheerders

### Setup voor Gedeeld Gebruik:
1. Voer `network_setup.bat` uit
2. Voer het gewenste netwerkpad in
3. Het systeem wordt automatisch gekopieerd

### Backup maken:
```bash
git add .
git commit -m "Backup voor wijzigingen"
```

### Configuratie herstellen:
- Kopieer `category_config.json` en `item_links.json` terug
- Of gebruik `git checkout` om terug te gaan 