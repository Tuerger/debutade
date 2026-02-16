# Contributie Debutade - Web Applicatie

Een web-applicatie die contributiebetalingen koppelt aan het ledenbestand.

## 📋 Overzicht

Deze applicatie leest ledengegevens uit **Ledenbestand.xlsx** en matcht die met banktransacties in het bankrekeningbestand. Het resultaat is een overzicht met ID, achternaam, rekening en het totaal betaalde bedrag.

## ✨ Functionaliteiten

- ✅ Leest ledengegevens uit tab **personen** en **betaald**
- ✅ Filtert banktransacties op tags **contributie-volwassenen** en **contributie-jeugd**
- ✅ Overzicht met totalen per tegenrekening
- ✅ Zelfde layout/stijl als de andere Debutade-apps
- ✅ Logging en duidelijke foutmeldingen

## 🚀 Installatie

### Vereisten

- Python 3.8 of hoger
- pip

### Installeren

```powershell
pip install -r requirements.txt
```

### Starten

```powershell
$env:DEBUTADE_CONFIG="C:\Debutade\config.json"
python webapp.py
```

De applicatie start op: **http://127.0.0.1:5004**

## 🔧 Configuratie

De configuratie staat in `c:\Debutade\config.json` onder de sectie `contributie`.

Voorbeeld:

```json
{
    "contributie": {
        "ledenbestand_path": "C:\\Users\\ericg\\OneDrive - Vereniging met volledige rechtsbevoegdheid\\SharePoint Debutade - Documenten\\03. Secretaris\\ledenadministratie\\Ledenbestand.xlsx",
        "leden_sheet_personen": "personen",
        "leden_sheet_betaald": "betaald",
        "bank_excel_file_name": "Debutade boekjaar bank 2026.xlsx",
        "bank_sheet_name": "Bankrekening",
        "tags": ["contributie-volwassenen", "contributie-jeugd"]
    }
}
```

De `bank_excel_file_name` wordt gecombineerd met de gedeelde `grootboek_directory` uit de `shared` sectie.

## 📊 Excel vereisten

**Ledenbestand.xlsx**
- Tab `personen`: kolommen `ID-lid` en `Achternaam`
- Tab `betaald`: kolommen `ID` en `Rekening`

**Bankrekening Excel**
- Tab `Bankrekening`
- Kolommen: `Tegenrekening`, `Bedrag (EUR)`, `Af Bij`, `Tag`

## 🧭 Gebruik

1. Open de webpagina
2. Controleer eventuele fouten bovenaan
3. Bekijk het overzicht met ID, achternaam, rekening en totaal betaald

## 🧪 Troubleshooting

- **Fout: Excel bestand niet gevonden**
  - Controleer de paden in `config.json`

- **Fout: Tabblad niet gevonden**
  - Controleer de tabnamen in de configuratie
