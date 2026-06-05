# Debutade Rapporten - Webapp

Deze webapp toont gecombineerde rapportage van **kasboek** en **bankrekening/spaarrekeningen**.

## Functionaliteit

- 4 saldikaarten bovenin:
  - Spaarrekening 1
  - Spaarrekening 2
  - Kas
  - Bankrekening
- Filters:
  - Maand
  - Tag
  - Bron (Kas, Bankrekening, Spaarrekening 1, Spaarrekening 2)
  - Vrije tekstzoeking in **mededelingen**
- Grafieken:
  - Verticale bar chart per **tag** met totaal positief en totaal negatief
  - Horizontale bar chart per **maand** met totaal positief en totaal negatief
- Tabel met transacties onderaan
- Export van de zichtbare tabel naar een **Word-document** (`.docx`) in `shared.grootboek_directory`
- Alle onderdelen reageren op dezelfde filters (kaarten, grafieken en tabel)

## Databronnen

De app leest transacties uit:

- Kasbestand uit `config.json` (`kasboek.excel_file_name` + `shared.grootboek_directory`)
- Bankbestand uit `config.json` (`bankrekening` / `shared.bank_excel_file_name` + `shared.grootboek_directory`)
- Standaard sheets:
  - Kas: `Kas`
  - Bank: `Bankrekening`, `Spaarrekening 1`, `Spaarrekening 2`

Optioneel kun je in `config.json` een sectie `rapporten` toevoegen om namen/sheets te overriden.

## Installatie

Installeer dependencies in de Python-omgeving van dit project:

```powershell
pip install -r project-debutade-rapporten/requirements.txt
```

## Starten

### Via hoofdapp (aanbevolen)

- Open Debutade Start
- Klik op kaart **Rapporten**

### Los starten

```powershell
cd project-debutade-rapporten
python webapp.py
```

Standaard draait de app op de poort uit `DEBUTADE_APP_PORT` (fallback: `5004`).

## API

De frontend haalt data op via:

- `GET /api/report-data`
- `POST /api/export-table-docx`

Response bevat:

- `transactions`: genormaliseerde transacties
- `filters`: beschikbare maanden/tags/bronnen
- `warnings`: waarschuwingen (bijv. als Excel-bestand ontbreekt)

De exportroute schrijft een Word-document weg in de map uit `shared.grootboek_directory` met een bestandsnaam in de vorm `transacties debutade <timestamp>.docx`.

## Bestanden

- `webapp.py` - Flask backend + Excel parsing + API
- `templates/index.html` - Dashboard UI met filters, grafieken en tabel
- `requirements.txt` - Python dependencies
