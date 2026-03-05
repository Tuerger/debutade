# Transactie Toevoegen - Debutade

Web applicatie voor het toevoegen van nieuwe bankrekening transacties vanuit een CSV bestand.

## Features

- **Filteren op sheet**: Selecteer tussen Bankrekening, Spaarrekening 1, of Spaarrekening 2
- **CSV Upload**: Laad transacties uit een CSV bestand
- **Data Profiling**: Validatie van datum, bedrag, af/bij en rekeningnummers
- **Duplicate Detection**: Controle op duplicaten voordat toevoegen
- **Preview**: Preview van transacties voordat ze daadwerkelijk worden toegevoegd
- **Backup**: Automatische backup van Excel bestand bij toevoegen
- **Sheet Statistieken**: Aantal records en datum van laatste transactie per sheet

## ESV Bestand Formaat

Het CSV bestand moet de volgende kolommen bevatten (hoofdlettergevoelig):
- `Datum` - Datum in format DD-MM-YYYY, YYYY-MM-DD, etc.
- `Naam / Omschrijving` (of `Omschrijving`, `naam_omschrijving`)
- `Af Bij` - Waarde: "Af" of "Bij"
- `Bedrag` - Numerieke waarde met punt of komma als decimaal separator
- `Rekening` - Rekeningnummer (geverifieerd per sheet)
- `Valutadatum` (optioneel)
- `Mededelingen` (optioneel)

## Validatieregels

- **Datum**: Moet geldig zijn en wordt geconverteerd naar Excel formaat (DD-MM-YYYY)
- **Bedrag**: Moet numeriek zijn
- **Af Bij**: Mag alleen "Af" of "Bij" zijn
- **Rekeningnummers per sheet**:
  - Bankrekening: `NL40INGB0002691632`
  - Spaarrekening 1: `S 858-17363`
  - Spaarrekening 2: `D 130-41072`
- **Duplicaten**: Gebaseerd op datum, bedrag en omschrijving

## Opzetten

1. Installeer dependencies: `pip install -r requirements.txt`
2. Zorg dat config.json correct is ingesteld met bankrekening bestandspaden
3. Start app: `python webapp.py`
4. Open browser op http://127.0.0.1:5005

## Configuratie

Zorg ervoor dat volgende instellingen in config.json correct zijn:
```json
{
    "shared": {
        "bank_excel_file_name": "Debutade boekjaar bank 2026.xlsx",
        "backup_directory": "C:\\path\\to\\backup",
        "log_directory": "C:\\path\\to\\logs"
    },
    "bankrekening": {
        "required_sheets": ["Bankrekening", "Spaarrekening 1", "Spaarrekening 2"]
    }
}
```

## Poort

Default draait de app op poort 5005. Dit kan worden overschreven met environment variable:
```bash
set DEBUTADE_APP_PORT=5006
```
