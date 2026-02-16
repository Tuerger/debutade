# Release Notes

## Version 1.0.0 - Initial Release (2026-01-22)

### 🎯 Overview

Eerste release van de "Voeg Bon Toe" web applicatie voor het koppelen van SharePoint bon URLs aan kasboek en bankrekening transacties.

### ✨ Features

- **Multi-bestand ondersteuning**: Werk simultaan met kasboek én bankrekening Excel bestanden
- **Multi-tab weergave**: Toon alle tabs uit beide Excel bestanden in één overzicht
- **SharePoint integratie**: 
  - Validatie dat alleen SharePoint URLs worden geaccepteerd
  - Ondersteuning voor PDF en JPG/JPEG bonnen
  - Sleep-en-plak functionaliteit
- **Automatische backups**: Bij elke server start worden backups gemaakt van beide Excel bestanden
- **Achtergrond opslaan**: URL opslag gebeurt asynchroon zonder de UI te blokkeren
- **Statistieken per tab**: Overzicht van records met/zonder bonnen
- **Responsive web interface**: Werkt op desktop, tablet en mobiel
- **Uitgebreide logging**: Alle acties worden gelogd met timestamps

### 🔧 Technical Details

- **Framework**: Flask (Python web framework)
- **Excel verwerking**: openpyxl voor lezen/schrijven van .xlsx bestanden
- **Threading**: ThreadPoolExecutor (4 workers) voor asynchroon opslaan
- **Caching**: Workbook caching met mtime-based invalidation voor performance
- **Port**: 5002 (localhost only)

### 📋 Configuration

Configuratie via `c:\Debutade\config.json` (sectie `bontoevoegen`):
```json
{
  "bontoevoegen": {
    "bank_excel_file_path": "pad/naar/bank.xlsx",
    "kas_excel_file_path": "pad/naar/kas.xlsx",
    "backup_directory": "pad/naar/backup",
    "log_directory": "pad/naar/log",
    "log_level": "INFO"
  }
}
```

### 🚀 Performance Optimizations

- Workbook caching in geheugen voor snellere toegang
- Asynchroon opslaan via thread pool (non-blocking UI)
- Smart reloading: alleen bij gewijzigde mtime
- Concurrent file access protection via threading locks

### 📊 File Requirements

Excel bestanden moeten bevatten:
- Kolom "Bon" voor opslag van URLs
- Kolom "Datum" voor transactiedatum
- Kolom "Naam / Omschrijving" voor beschrijving
- Kolom "Bedrag (EUR)" voor bedrag

### 🐛 Known Limitations

1. Alleen lokaal gebruik (localhost:5002)
2. Geen multi-user concurrency protection
3. Alleen SharePoint URLs toegestaan (by design)
4. Excel bestanden moeten .xlsx formaat zijn (geen .xls)

### 📚 Documentation

- [README_WEBAPP.md](README_WEBAPP.md) - Volledige gebruikersdocumentatie

---

## Version 1.1.0 - 2026-01-23

### 🆕 Belangrijkste wijzigingen

- **SharePoint tenant validatie**: Alleen URLs van de ingestelde tenant (configurabel via instellingen) worden geaccepteerd.
- **SharePoint tenant instelbaar**: Tenant-naam kan nu via de instellingenpagina worden gewijzigd.
- **Verbeterde URL-validatie**: Alleen SharePoint URLs met het juiste formaat en tenant worden geaccepteerd. Duidelijke foutmeldingen bij ongeldige URLs.
- **Records sortering**: Alle transacties worden nu standaard van nieuw (boven) naar oud (beneden) getoond, ongeacht tab.
- **Datumweergave**: Alleen dag-maand-jaar wordt getoond in de tabel, geen tijd.
- **Knopkleur en tekst**: "Bewaar" (groen) voor nieuwe bon, "Vervang" (blauw) als er al een bon-URL is.
- **Invoerveld leegmaken**: Bij foutieve of ongeldige URL wordt het plakveld automatisch geleegd.
- **Foutmeldingen**: Alle foutmeldingen zijn nu in het Nederlands en tonen de specifieke oorzaak (zoals verkeerde tenant).
- **Settings UI**: SharePoint tenant-naam toegevoegd aan instellingenpagina.

### 🐞 Overige verbeteringen

- Diverse kleine bugfixes en UI-verbeteringen.
- Code opgeschoond en logging verbeterd.

---
