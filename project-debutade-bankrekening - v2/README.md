
# Bankrekening Debutade - Web Applicatie
Een moderne web-gebaseerde applicatie voor het beheren van bankrekeningtransacties van Debutade.

## 📋 Overzicht

Deze web applicatie is een modernisering van de originele Tkinter desktop applicatie. Het biedt dezelfde functionaliteit via een gebruiksvriendelijke webinterface die toegankelijk is via elke moderne webbrowser.

## ✨ Functionaliteiten

- ✅ **AI Tag Recommender**: Automatische tag-suggesties op basis van trainingsdata
- ✅ **Backup functie**: Maak handmatig of automatisch backups
- ✅ **Logging**: Uitgebreide logging van alle acties
- ✅ **Tags**: Voeg tags (categorieen) toe aan transacties
- ✅ **Responsive design**: Werkt op desktop, tablet en mobiel

## 🚀 Installatie

### Vereisten

- **Python 3.12** of hoger (Python 3.13 heeft compatibiliteitsproblemen met scikit-learn)
- pip (Python package manager)

### Stap 1: Clone of download de bestanden

Zorg dat je de volgende bestanden hebt:
```
bankrekening_debutade/code/
├── webapp.py
├── tag_recommender.py     # AI module voor tag suggesties
├── requirements.txt
├── templates/
│   ├── base.html
│   └── index.html
└── static/
    ├── style.css
    └── category_test_set.xlsx  # Trainingsdata (niet in git)
```

### Stap 2: Maak een Python 3.12 virtual environment

```powershell
# Maak venv aan (voor Python 3.12)
py -3.12 -m venv .venv312

# Activeer venv (Windows)
.\.venv312\Scripts\Activate.ps1

# Of (Mac/Linux)
source .venv312/bin/activate
```

### Stap 3: Installeer dependencies

```powershell
python -m pip install --upgrade pip
pip install -r requirements.txt
```

### Stap 4: Configuratie

De configuratie staat centraal in het bestand `c:\Debutade\config.json` en wordt beheerd via de hoofdapp (menu **Instellingen**).

Voorbeeld van de `bankrekening`-sectie in `config.json`:

```json
{
    "bankrekening": {
        "excel_file_path": "C:\\pad\\naar\\Debutade boekjaar bank 2026.xlsx",
        "backup_directory": "C:\\pad\\naar\\backup",
        "resources": "C:\\pad\\naar\\resources",
        "log_directory": "C:\\pad\\naar\\log",
        "excel_sheet_name": "Bankrekening",
        "required_sheets": [
            "Bankrekening",
            "Spaarrekening 1",
            "Spaarrekening 2"
        ],
        "tags": ["500;Vermogen Debutade", "510;Reserveringen"],
        "log_level": "INFO"
    }
}
```

**Let op**: Zorg dat de opgegeven directories bestaan of dat de applicatie rechten heeft om ze aan te maken.

### Stap 5: Start de applicatie

#### Optie 1: Start via de hoofdapp

Gebruik de Debutade hoofdapp en start daar de bankrekening-app.

#### Optie 2: Standalone (met centraal configbestand)

```powershell
$env:DEBUTADE_CONFIG="C:\Debutade\config.json"
python webapp.py
```

De applicatie start op: **http://127.0.0.1:5001**

## 💻 Gebruik

### Transactie toevoegen

1. Open je browser en ga naar `http://127.0.0.1:5001`
2. Vul het formulier in:
   - **Datum**: Selecteer de transactiedatum
   - **Naam/Omschrijving**: Beschrijving van de transactie (verplicht)
   - **Af/Bij**: Kies of geld uit de kas gaat (Af) of erin komt (Bij)
   - **Bedrag**: Voer het bedrag in (gebruik komma of punt als decimaal)
   - **Mutatiesoort**: Standaard "Kas"
   - **Tag**: Optioneel - categoriseer de transactie
3. Klik op **Opslaan**
4. De transactie wordt toegevoegd en het banksaldo wordt bijgewerkt

### AI Tag Suggesties gebruiken

1. Scroll naar "Transacties zonder Tag" op de hoofdpagina
2. Klik op **AI suggestie** naast een transactie
3. Het systeem haalt de top 3 tag-suggesties op:
   - De beste suggestie wordt automatisch ingevuld
   - Alle suggesties worden getoond met hun scores
4. Klik op een suggestie-chip om een andere te kiezen
5. Klik op **Opslaan** om de tag op te slaan

**Let op**: De AI module werkt alleen als `static/category_test_set.xlsx` bestaat en trainingsdata bevat. Zie [README_AI_MODULE.md](README_AI_MODULE.md) voor meer details.

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
├── tag_recommender.py     # AI module voor tag suggesties
├── requirements.txt       # Python dependencies
├── templates/             # HTML templates
│   ├── base.html         # Basis template
│   └── index.html        # Hoofdpagina
└── static/               # Statische bestanden
    ├── style.css         # Custom CSS styling
    └── category_test_set.xlsx  # Trainingsdata (niet in git)
```

## 🔧 Configuratie opties

| Optie | Beschrijving |
|-------|-------------|
| `excel_file_path` | Volledig pad naar het Excel bestand |
| `backup_directory` | Map voor backup bestanden |
| `resources` | Map met resource-bestanden |
| `log_directory` | Map voor log bestanden |
| `excel_sheet_name` | Naam van het Excel sheet/tabblad |
| `required_sheets` | Vereiste sheetnamen in het Excel bestand |
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

Vereiste tabs (sheets):
- Bankrekening
- Spaarrekening 1
- Spaarrekening 2

Alle drie de tabs moeten exact bovenstaande kolomheaders bevatten (in dezelfde volgorde en schrijfwijze).

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

### Fout: "Excel bestand voldoet niet aan het vereiste formaat"
- Het bestand moet exact 3 tabs bevatten: Bankrekening, Spaarrekening 1, Spaarrekening 2
- Elke tab moet de exacte kolomheaders hebben: "Datum; Naam / Omschrijving; Rekening; Tegenrekening; Code; Af Bij; Bedrag (EUR); Mutatiesoort; Mededelingen; Saldo na mutatie; ; Tag"
- Pas de namen en headers aan in jouw Excel of kies een ander bestand

### Instellen van Sheet-naam
- De sheet-naam kan alleen één van de drie vereiste tab-namen zijn
- Bij een andere naam geeft de applicatie een duidelijke foutmelding

### Applicatie start niet
- Controleer of alle dependencies zijn geïnstalleerd: `pip install -r requirements.txt`
- Controleer of poort 5001 niet al in gebruik is

### Locale waarschuwing
- Dit is normaal op systemen zonder Nederlandse locale
- De applicatie blijft gewoon werken

### AI suggesties werken niet
- Controleer of `static/category_test_set.xlsx` bestaat
- Het bestand moet een "Tag" kolom bevatten
- Er moeten minimaal 10-20 trainingsvoorbeelden per tag zijn voor goede resultaten
- Zie [README_AI_MODULE.md](README_AI_MODULE.md) voor gedetailleerde troubleshooting

## 📝 Logging

Alle acties worden gelogd in: `{log_directory}/bankrekening_webapp_log.txt`

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
4. Voor AI module problemen, zie [README_AI_MODULE.md](README_AI_MODULE.md)

## 📄 Licentie

© 2026 Debutade - Voor intern gebruik

## 👤 Auteur

Eric G.

---

**Versie**: 2.1 (Web App + AI Module)  
**Datum**: 2026-01-07  
**Gebaseerd op**: bankrekening_debutade.py v1.0

## 📚 Aanvullende Documentatie

- [README_AI_MODULE.md](README_AI_MODULE.md) - Uitgebreide documentatie over de AI tag recommender
- [CHANGES_AI_MODULE.md](CHANGES_AI_MODULE.md) - Overzicht van wijzigingen voor de AI module
