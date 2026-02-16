# Voeg Bon Toe - Web Applicatie

Een web-gebaseerde applicatie voor het toevoegen van bon URLs aan kasboek en bankrekening records van Debutade.

## 📋 Overzicht

Deze applicatie maakt het eenvoudig om SharePoint links naar bonnetjes (PDF/JPG) te koppelen aan financiële transacties in Excel bestanden. De applicatie toont alle records uit zowel het kasboek als de bankrekening, en biedt per record de mogelijkheid om een bon URL toe te voegen.

## ✨ Functionaliteiten

- ✅ **Multi-bestand ondersteuning**: Werk met kasboek én bankrekening Excel bestanden tegelijk
- ✅ **Multi-tab weergave**: Toon alle tabs uit beide Excel bestanden in één overzicht
- ✅ **SharePoint integratie**: Sleep-en-plak functionaliteit voor SharePoint bon URLs
- ✅ **Automatische backup**: Bij elke start worden backups van beide bestanden gemaakt
- ✅ **Achtergrond opslaan**: URLs worden asynchroon opgeslagen zonder de UI te blokkeren
- ✅ **Statistieken**: Per tab overzicht van hoeveel bonnen aanwezig zijn
- ✅ **Logging**: Uitgebreide logging van alle acties
- ✅ **Responsive design**: Werkt op desktop, tablet en mobiel

## 🚀 Installatie

### Vereisten

- Python 3.8 of hoger
- pip (Python package manager)

### Stap 1: Clone of download de bestanden

Zorg dat je de volgende bestanden hebt:
```
project-debutade-bontoevoegen/
├── voegbontoe_webapp.py          # Hoofdapplicatie
├── start_voegbontoe.py           # Start script
├── Start Voegbontoe.bat          # Windows start bestand
├── requirements.txt              # Python dependencies
├── templates/
│   └── voegbontoe.html          # HTML template
└── static/                       # (optioneel voor CSS/JS)
```

### Stap 2: Installeer dependencies

Open een terminal/PowerShell in de project directory en voer uit:

```powershell
pip install -r requirements.txt
```

### Stap 3: Configuratie

De configuratie staat centraal in het bestand `c:\Debutade\config.json` en wordt beheerd via de hoofdapp (menu **Instellingen**).

Voorbeeld van de `bontoevoegen`-sectie in `config.json`:

```json
{
    "bontoevoegen": {
        "bank_excel_file_path": "C:\\pad\\naar\\Debutade boekjaar bank 2026.xlsx",
        "kas_excel_file_path": "C:\\pad\\naar\\Debutade boekjaar kas 2026.xlsx",
        "backup_directory": "C:\\pad\\naar\\backup",
        "log_directory": "C:\\pad\\naar\\log",
        "log_level": "INFO",
        "sharepoint_tenant": "debutabe"
    }
}
```

**Let op**: Zorg dat de opgegeven Excel bestanden bestaan of dat de applicatie rechten heeft om ze te lezen/schrijven.

### Stap 4: Start de applicatie

#### Optie 1: Gebruik Windows batch bestand

Dubbelklik op `Start Voegbontoe.bat`

#### Optie 2: Via Python

```powershell
$env:DEBUTADE_CONFIG="C:\Debutade\config.json"
python voegbontoe_webapp.py
```

#### Optie 3: Via start script

```powershell
python start_voegbontoe.py
```

De applicatie start op: **http://127.0.0.1:5002**

## 💻 Gebruik

### Bonnen toevoegen aan records

1. Open je browser en ga naar `http://127.0.0.1:5002`
2. Je ziet een overzicht van alle records uit beide Excel bestanden (kasboek + bankrekening)
3. Records worden gegroepeerd per tab en bestand
4. Voor elke record zie je:
   - Datum
   - Naam/Omschrijving
   - Bedrag
   - Bon status (✓ als er al een bon gekoppeld is)
   - Tab en bron (Kasboek of Bankrekening)

### Bon URL toevoegen

Er zijn twee manieren om een bon URL toe te voegen:

#### Methode 1: Sleep en plak (aanbevolen)
1. Ga naar SharePoint en zoek het bonnetje (PDF of JPG)
2. Kopieer de SharePoint link
3. Sleep de link naar het "Bon URL" invoerveld van de juiste record
4. Klik op **Bewaar**

#### Methode 2: Handmatig invoeren
1. Kopieer de SharePoint URL van het bonnetje
2. Plak de URL in het "Bon URL" invoerveld
3. Klik op **Bewaar**

**Let op**: 
- Alleen SharePoint URLs worden geaccepteerd (beveiliging)
- Het opslaan gebeurt asynchroon op de achtergrond
- Wacht enkele seconden en ververs de pagina om te zien of de bon is opgeslagen

### Statistieken bekijken

Bovenaan de pagina zie je per tab:
- Totaal aantal records
- Aantal records mét bon (✓)
- Aantal records zonder bon

### Applicatie afsluiten

Klik op **Afsluiten** in de navigatiebalk om de server veilig af te sluiten.

## 📁 Bestandsstructuur

```
project-debutade-bontoevoegen/
├── voegbontoe_webapp.py          # Flask web applicatie
├── start_voegbontoe.py           # Alternatief start script
├── Start Voegbontoe.bat          # Windows batch start bestand
├── requirements.txt              # Python dependencies (Flask, openpyxl)
├── templates/
│   └── voegbontoe.html          # HTML template
├── static/                       # (optioneel) CSS/JS/images
├── backup/                       # Automatische backups (wordt aangemaakt)
└── logs/                         # Log bestanden (wordt aangemaakt)
```

## 🔧 Configuratie opties

| Optie | Beschrijving | Vereist |
|-------|-------------|---------|
| `bank_excel_file_path` | Pad naar het bankrekening Excel bestand | ✓ |
| `kas_excel_file_path` | Pad naar het kasboek Excel bestand | ✓ |
| `backup_directory` | Map voor backup bestanden | Optioneel |
| `log_directory` | Map voor log bestanden | Optioneel |
| `log_level` | Logniveau (DEBUG, INFO, WARNING, ERROR) | Optioneel |
| `sharepoint_tenant` | SharePoint tenant naam | Optioneel |

## 📊 Excel bestand vereisten

De Excel bestanden moeten de volgende kolommen bevatten:
- **Datum**: Datum van de transactie
- **Naam / Omschrijving**: Omschrijving van de transactie
- **Bedrag (EUR)**: Bedrag in euro's
- **Bon**: Kolom waar de SharePoint URL wordt opgeslagen (mag leeg zijn)

**Belangrijk**:
- Elke Excel bestand kan meerdere tabs bevatten
- Alle tabs worden in de applicatie getoond
- De kolom "Bon" moet aanwezig zijn in elke tab
- De eerste rij moet de headers bevatten

## 🔐 Beveiliging

**Let op**: Deze applicatie is bedoeld voor lokaal gebruik binnen een vertrouwde omgeving.

Beveiligingsmaatregelen:
- Alleen SharePoint URLs worden geaccepteerd (voorkomt willekeurige links)
- Applicatie draait standaard op localhost (127.0.0.1)
- Poort 5002 is alleen lokaal toegankelijk
- Geen externe netwerkverbindingen toegestaan

Voor productiegebruik (niet aanbevolen):
- Voeg authenticatie toe (bijv. Flask-Login)
- Gebruik HTTPS met SSL certificaten
- Configureer een productie-webserver (bijv. Gunicorn + Nginx)
- Implementeer rate limiting

## 🐛 Troubleshooting

### Fout: "Configuratiebestand niet gevonden"
- Controleer of `c:\Debutade\config.json` bestaat
- Controleer of het bestand correct geformatteerd is (geldige JSON)

### Fout: "Excel-bestand niet gevonden"
- Controleer of de paden in `c:\Debutade\config.json` correct zijn
- Gebruik absolute paden (bijv. `C:\\pad\\naar\\bestand.xlsx`)
- Zorg dat de bestanden niet door een ander programma zijn geopend

### Fout: "Kolom 'Bon' niet gevonden"
- Open het Excel bestand en controleer of elke tab een kolom met de header "Bon" heeft
- Let op hoofdletters: het moet exact "Bon" zijn

### Fout: "Poort 5002 is al in gebruik"
Oplossingen:
1. Sluit alle vorige "Voeg Bon Toe" vensters
2. Open PowerShell en voer uit: `Get-Process -Name python | Stop-Process -Force`
3. Of voer uit: `netstat -ano | findstr :5002` om het proces te vinden en te stoppen

### Bon URL wordt niet opgeslagen
- Wacht 5-10 seconden na het klikken op "Bewaar"
- Ververs de pagina (F5) om te controleren of de bon is opgeslagen
- Controleer het logbestand in de log directory voor foutmeldingen
- Zorg dat het Excel bestand niet alleen-lezen is

### Applicatie start niet
- Controleer of Python 3.8+ is geïnstalleerd: `python --version`
- Controleer of alle dependencies zijn geïnstalleerd: `pip install -r requirements.txt`
- Controleer de log bestanden voor foutmeldingen

## 📝 Logging

Alle acties worden gelogd in: `{log_directory}/voegbontoe_YYYYMMDD.log`

Log entries bevatten:
- Timestamp
- Log level (INFO, WARNING, ERROR)
- Actie/gebeurtenis
- Details over opgeslagen URLs en eventuele fouten

Voorbeeld log entries:
```
2026-01-23 10:30:00 - INFO - Bon URL opgeslagen (cache): path.xlsx, tab=Rekening 1, rij=5
2026-01-23 10:31:00 - WARNING - Excel bestand niet gevonden: path.xlsx
2026-01-23 10:32:00 - ERROR - Fout bij opslaan bon URL: Permission denied
```

## 🚀 Performance Features

- **Workbook caching**: Excel bestanden worden in geheugen gehouden voor snellere toegang
- **Asynchroon opslaan**: URL opslag gebeurt in achtergrond threads (niet-blokkerend)
- **Thread pool**: Gebruik van 4 worker threads voor parallelle verwerking
- **Smart reloading**: Workbooks worden alleen opnieuw geladen bij wijziging (mtime check)

Dit zorgt voor snelle response tijden, zelfs bij grote Excel bestanden.

## 🆘 Ondersteuning

Voor vragen of problemen:
1. Controleer de logbestanden in `{log_directory}`
2. Controleer de browserconsole (F12) voor JavaScript fouten
3. Zorg dat alle paden in `c:\Debutade\config.json` correct zijn
4. Controleer of de Excel bestanden de vereiste kolommen bevatten

## 📄 Licentie

© 2026 Debutade - Voor intern gebruik

## 👤 Auteur

Eric G.

---

**Versie**: 1.0  
**Datum**: 2026-01-22  
**Project**: Voeg Bon Toe Web Applicatie
