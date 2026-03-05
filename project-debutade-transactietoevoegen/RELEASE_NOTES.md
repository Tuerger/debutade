# RELEASE NOTES - Transactie Toevoegen

## Versie 1.0 - 2026-03-05

### Features

**Interface**
- Responsive web interface gebaseerd op Debutade design system
- Donkere theme met accent kleuren
- Real-time statistieken kaarten
- Sheet filter (Bankrekening, Spaarrekening 1, Spaarrekening 2)

**Transactie Management**
- CSV file import met flexible kolom matching
- Preview van transacties voordat toevoegen (cursief weergave)
- Automatische duplicate detection
- Batch toevoegen van transacties

**Data Profiling**
- Datum validatie en formaat conversie
- Bedrag numerieke validatie
- Af/Bij enum validatie
- Rekeningnummer validatie per sheet:
  - Bankrekening: `NL40INGB0002691632`
  - Spaarrekening 1: `S 858-17363`
  - Spaarrekening 2: `D 130-41072`

**Beveiliging & Data Integriteit**
- Automatische backup voor toevoegen
- Controle op duplicaten
- Foutafhandeling en validatie feedback
- Logging van alle operaties

**Statistieken**
- Aantal records per sheet
- Laatste transactie datum
- Real-time updates na toevoegen

### Configuratie

Ondersteunde config settings uit config.json:
```json
{
    "shared": {
        "bank_excel_file_name": "Debutade boekjaar bank 2026.xlsx",
        "backup_directory": "C:\\path\\to\\backup",
        "log_directory": "C:\\path\\to\\logs",
        "log_level": "INFO|DEBUG|WARNING|ERROR"
    },
    "bankrekening": {
        "required_sheets": ["Bankrekening", "Spaarrekening 1", "Spaarrekening 2"]
    }
}
```

### Gekende Beperkingen

- CSV kolommen moeten case-insensitive matchen met verwachte namen
- Excel bestand moet al bestaan met correct sheet structuur
- Saldo na mutatie kolom wordt niet automatisch ingevuld (kan handmatig)
- Enkel bankrekening sheets ondersteund (geen Kas)

### API Endpoints

- `GET /` - Main interface
- `GET /api/sheet-data/<sheet_name>` - Laad transacties voor sheet
- `POST /api/parse-csv` - Parse en valideer CSV bestand
- `POST /api/add-transactions` - Voeg gevalideerde transacties toe
- `POST /quit` - Sluit applicatie af

### Dependencies

- Flask 3.0.0+
- openpyxl 3.1.0+
- Python 3.8+

### Changelog

**v1.0 (2026-03-05)**
- Initial release
- Full CSV import met validatie
- Sheet filtering
- Duplicate detection
- Backup functionality
- Statistics dashboard

---

**Voor vragen of bugs**: Controleer het log bestand in config.json `log_directory`
