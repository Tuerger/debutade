# Kasboek Debutade - Web Applicatie

Een moderne web-gebaseerde applicatie voor het beheren van kasboektransacties van Debutade.

## 📋 Overzicht

Deze web applicatie is een modernisering van de originele Tkinter desktop applicatie. Het biedt dezelfde functionaliteit via een gebruiksvriendelijke webinterface die toegankelijk is via elke moderne webbrowser.

## ✨ Functionaliteiten

- ✅ **Transacties toevoegen**: Voeg nieuwe kas transacties toe via een intuïtief webformulier
- ✅ **Validatie**: Automatische controle van datums en bedragen
- ✅ **Excel integratie**: Automatische opslag in Excel-bestand (`records.xlsx`)
- ✅ **Recente transacties**: Overzicht van de laatste transacties in real-time
- ✅ **Saldo berekening**: Automatische berekening van het totale kassaldo
- ✅ **Backup functie**: Maak handmatig of automatisch backups
- ✅ **Logging**: Uitgebreide logging van alle acties
- ✅ **Tags**: Categoriseer transacties met tags
- ✅ **Responsive design**: Werkt op desktop, tablet en mobiel

## 🚀 Installatie

### Vereisten

- Python 3.8 of hoger
- pip (Python package manager)

### Stap 1: Clone of download de bestanden

Zorg dat je de volgende bestanden hebt:
```
kasboek_debutade/code/
├── webapp.py
├── requirements.txt
├── templates/
│   ├── base.html
│   └── index.html
└── static/
    └── style.css
```

### Stap 2: Installeer dependencies

Open een terminal/PowerShell in de code directory en voer uit:

```powershell
pip install -r requirements.txt
```

### Stap 3: Configuratie

De configuratie staat centraal in het bestand `c:\Debutade\config.json` en wordt beheerd via de hoofdapp (menu **Instellingen**).

Voorbeeld van de `kasboek`-sectie in `config.json`:

```json
{
    "kasboek": {
        "excel_file_path": "C:\\pad\\naar\\Debutade boekjaar kas 2026.xlsx",
        "backup_directory": "C:\\pad\\naar\\backup",
        "resources": "C:\\pad\\naar\\resources",
        "log_directory": "C:\\pad\\naar\\log",
        "excel_sheet_name": "Kas",
        "tags": ["500;Vermogen Debutade", "510;Reserveringen"],
        "log_level": "INFO"
    }
}
```

**Let op**: Zorg dat de opgegeven directories bestaan of dat de applicatie rechten heeft om ze aan te maken.

### Stap 4: Start de applicatie

#### Optie 1: Start via de hoofdapp

Gebruik de Debutade hoofdapp en start daar de kasboek-app.

#### Optie 2: Standalone (met centraal configbestand)

```powershell
$env:DEBUTADE_CONFIG="C:\Debutade\config.json"
python webapp.py
```

De applicatie start op: **http://127.0.0.1:5000**

## 💻 Gebruik

### Transactie toevoegen

1. Open je browser en ga naar `http://127.0.0.1:5000`
2. Vul het formulier in:
   - **Datum**: Selecteer de transactiedatum
   - **Naam/Omschrijving**: Beschrijving van de transactie (verplicht)
   - **Af/Bij**: Kies of geld uit de kas gaat (Af) of erin komt (Bij)
   - **Bedrag**: Voer het bedrag in (gebruik komma of punt als decimaal)
   - **Mutatiesoort**: Standaard "Kas"
   - **Tag**: Optioneel - categoriseer de transactie
3. Klik op **Opslaan**
4. De transactie wordt toegevoegd en het kassaldo wordt bijgewerkt

### Recente transacties bekijken

De rechterkolom toont automatisch de 10 meest recente transacties. Deze lijst wordt elke 30 seconden automatisch ververst.

### Instellingen bekijken

Instellingen zijn alleen beschikbaar via de hoofdapp (menu **Instellingen**).

### Backup maken

- Automatisch: Bij elke start van de applicatie wordt een backup gemaakt
- Handmatig: Klik op **Backup** in de navigatiebalk

## 📁 Bestandsstructuur

```
code/
├── webapp.py              # Hoofdapplicatie (Flask)
├── requirements.txt       # Python dependencies
├── templates/             # HTML templates
│   ├── base.html         # Basis template
│   └── index.html        # Hoofdpagina
└── static/               # Statische bestanden
    └── style.css         # Custom CSS styling
```

## 🔧 Configuratie opties

| Optie | Beschrijving |
|-------|-------------|
| `excel_file_path` | Volledig pad naar het Excel bestand |
| `backup_directory` | Map voor backup bestanden |
| `resources` | Map met resource-bestanden |
| `log_directory` | Map voor log bestanden |
| `excel_sheet_name` | Naam van het Excel sheet/tabblad |
| `tags` | Lijst van beschikbare tags |
| `log_level` | Logniveau (DEBUG, INFO, WARNING, ERROR) |

## 📊 Excel bestand structuur

Het Excel bestand heeft de volgende kolommen:
1. Datum
2. Naam/Omschrijving
3. Rekening
4. Tegen Rekening
5. Code
6. Af Bij
7. Bedrag
8. Mutatiesoort
9. Mededelingen
10. Saldo na mutatie
11. (leeg)
12. Tag

## 🔐 Beveiliging

**Let op**: Deze applicatie is bedoeld voor lokaal gebruik. Voor productiegebruik:
- Zet `debug=False` in `webapp.py`
- Voeg authenticatie toe
- Gebruik HTTPS
- Configureer een productie-webserver (bijv. Gunicorn + Nginx)

## 🐛 Troubleshooting

### Fout: "Configuratiebestand niet gevonden"
- Controleer of `c:\Debutade\config.json` bestaat
- Gebruik de omgevingsvariabele `DEBUTADE_CONFIG` om het pad op te geven

### Fout: "Excel-bestand niet gevonden"
- Zorg dat het Excel bestand bestaat op het opgegeven pad
- Of laat de applicatie een nieuw bestand aanmaken door een transactie toe te voegen

### Applicatie start niet
- Controleer of alle dependencies zijn geïnstalleerd: `pip install -r requirements.txt`
- Controleer of poort 5000 niet al in gebruik is

### Locale waarschuwing
- Dit is normaal op systemen zonder Nederlandse locale
- De applicatie blijft gewoon werken

## 📝 Logging

Alle acties worden gelogd in: `{log_directory}/kasboek_webapp_log.txt`

Log entries bevatten:
- Timestamp
- Log level (INFO, WARNING, ERROR)
- Actie/gebeurtenis
- IP adres van de gebruiker (bij transacties)

## 🔄 Verschillen met desktop versie

| Feature | Desktop (Tkinter) | Web App |
|---------|------------------|----------|
| Interface | Desktop venster | Webbrowser |
| Toegang | Lokale machine | Lokaal netwerk mogelijk |
| Styling | Tkinter widgets | Modern Bootstrap design |
| Real-time updates | N/A | Auto-refresh transacties |
| Multi-user | Nee | Mogelijk (met voorzichtigheid) |

## 🆘 Ondersteuning

Voor vragen of problemen:
1. Controleer de logbestanden in `{log_directory}`
2. Controleer de browserconsole (F12) voor JavaScript fouten
3. Zorg dat alle paden in `c:\Debutade\config.json` correct zijn

## 📄 Licentie

© 2026 Debutade - Voor intern gebruik

## 👤 Auteur

Eric G.

---

**Versie**: 2.0 (Web App)  
**Datum**: 2026-01-03  
**Gebaseerd op**: kasboek_debutade.py v1.0
