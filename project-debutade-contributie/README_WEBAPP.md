# Contributie Debutade - Web Applicatie

Een web-applicatie die contributiebetalingen koppelt aan het ledenbestand via `ID-lid` in de bankmededelingen.

## 📋 Overzicht

Deze applicatie leest ledengegevens uit **Ledenbestand.xlsx** (tab `leden`) en zoekt per lid op `ID-lid` in de kolom `mededelingen` van het bankrekeningbestand. Het resultaat is een overzicht met:

- `ID-lid`
- `Achternaam`
- `Email`
- `Te innen bedrag`
- `Ontvangen bedrag`
- `Opmerking`
- `Status` (`✅`, `🔵`, `❌`)

## ✨ Functionaliteiten

- ✅ Leest ledengegevens uit tab **leden**
- ✅ Zoekt `ID-lid` in banktab **bankrekening** kolom **mededelingen**
- ✅ Berekent per lid ontvangen bedrag uit kolom **bedrag**
- ✅ Status per lid: volledig / gedeeltelijk / nog niets
- ✅ Handmatig als betaald markeren per lid met reden
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
        "leden_sheet_name": "leden",
        "bank_excel_file_name": "Debutade boekjaar 2026 Bank.xlsx",
        "bank_sheet_name": "bankrekening"
    }
}
```

De `bank_excel_file_name` wordt gecombineerd met de gedeelde `grootboek_directory` uit de `shared` sectie.

## 📊 Excel vereisten

**Ledenbestand.xlsx**
- Tab `leden`: kolommen `ID-lid`, `Achternaam`, `Email`, `bedrag`

**Bankrekening Excel**
- Tab `bankrekening`
- Kolommen: `mededelingen`, `bedrag` (optioneel ook `Af Bij`)

## 🧭 Gebruik

1. Open de webpagina
2. Controleer eventuele fouten bovenaan
3. Bekijk het overzicht met te innen bedrag, ontvangen bedrag en status per lid
4. Gebruik indien nodig de knop **Handmatig betaald** bij een lid en vul een reden in

### Handmatige betaald-markering

Een handmatige betaald-markering wordt opgeslagen in `config.json` onder:

```json
"contributie": {
  "manual_paid_overrides": {
    "<lidnummer>": {
      "marked_paid": true,
      "reason": "Betaald in vorig boekjaar",
      "updated_at": "2026-03-09 21:10:00"
    }
  }
}
```

Je kunt deze markering ook weer verwijderen vanuit hetzelfde overzicht.

## 🧪 Troubleshooting

- **Fout: Excel bestand niet gevonden**
  - Controleer de paden in `config.json`

- **Fout: Tabblad niet gevonden**
  - Controleer de tabnamen in de configuratie
