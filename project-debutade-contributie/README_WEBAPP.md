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

De applicatie-instellingen staan in `c:\Debutade\config.json` onder de sectie `contributie`. Leden- en betaalgegevens worden apart opgeslagen in `c:\Debutade\leden.json`.

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

Bij het opstarten worden bestaande dynamische ledengegevens uit `config.json` eenmalig naar `leden.json` verplaatst. Dit omvat ledenrecords, betaalstatussen, handmatige correcties en verwerkte transactiebesluiten. Daarna worden deze gegevens alleen in `leden.json` bijgehouden; `config.json` bevat de vaste applicatie-instellingen.

### Opbouw van `leden.json`

Alle dynamische informatie die bij een specifiek lid hoort, staat gebundeld onder `leden.json` -> `leden` -> `"<lidnummer>"`. Alleen informatie die niet aan één lid gekoppeld kan worden (zoals splitsingsbesluiten voor transacties met meerdere lidnummers) staat in de generieke sectie `transactie_afspraken`.

```json
{
  "leden": {
    "<lidnummer>": {
      "achternaam": "...",
      "email": "...",
      "rekeningnummer": "...",
      "due_amount": 290.0,
      "manual_transaction_mapping": "... (optioneel)",
      "manual_paid_override": { "marked_paid": true, "reason": "...", "updated_at": "..." },
      "manual_refund_override": { "amount": 0.0, "reason": "...", "updated_at": "..." },
      "opgezegd": true,
      "opgezegd_achternaam": "... (alleen als opgezegd)",
      "opgezegd_roepnaam": "... (alleen als opgezegd)",
      "opgezegd_email": "... (alleen als opgezegd)",
      "status": { "due_amount": 0.0, "received_amount": 0.0, "refunded_amount": 0.0, "status_label": "...", "status_class": "...", "updated_at": "..." },
      "terugstortingen": [
        { "mededelingen": "...", "amount": 145.0, "reason": "...", "processed_at": "..." }
      ]
    }
  },
  "transactie_afspraken": {
    "split_beslissingen": {
      "<mededelingen-tekst>": { "mode": "split | single_member", "member_id": "... (bij single_member)" }
    },
    "niet_gekoppelde_terugstortingen": [
      { "mededelingen": "...", "amount": 0.0, "reason": "...", "processed_at": "..." }
    ]
  }
}
```

Leden-basisgegevens (`achternaam`, `email`, `rekeningnummer`, `due_amount`) worden automatisch gesynchroniseerd vanuit tabblad `personen`, opzeggingsgegevens vanuit tabblad `opgezegd`. Handmatige velden, status en terugstortingen blijven per lid behouden, ook als een lid (tijdelijk) niet meer in het ledenbestand voorkomt. Wanneer het penningmeester-overzicht een "niet-gekoppelde" terugstorting verwerkt, probeert de applicatie deze automatisch aan een lid te koppelen op basis van het rekeningnummer in de banktekst; lukt dat niet, dan komt de terugstorting in `transactie_afspraken.niet_gekoppelde_terugstortingen` terecht.

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

Een handmatige betaald-markering wordt opgeslagen in `leden.json`, genest bij het betreffende lid onder `leden`:

```json
{
  "leden": {
    "<lidnummer>": {
      "manual_paid_override": {
        "marked_paid": true,
        "reason": "Betaald in vorig boekjaar",
        "updated_at": "2026-03-09 21:10:00"
      }
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
