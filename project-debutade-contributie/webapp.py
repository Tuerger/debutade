"""
Contributie Debutade - Web Applicatie
====================================

Web applicatie voor het koppelen van contributiebetalingen aan leden.

Functionaliteiten:
- Leest ledengegevens uit Ledenbestand.xlsx (tab personen en tab betaald)
- Leest banktransacties met tags contributie-volwassenen/jeugd
- Maakt een overzicht met ID, achternaam, rekening en totaal betaald

Versie: 1.0
Datum: 2026-02-16
Auteur: Eric G.
"""

from flask import Flask, render_template, jsonify
from openpyxl import load_workbook
from datetime import datetime
import logging
import os
import json
import sys
import re
import tempfile
import shutil

# Fix encoding voor Windows console
if sys.platform == "win32":
    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except AttributeError:
        import io
        sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")

app = Flask(__name__)
app.config["TEMPLATES_AUTO_RELOAD"] = True

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
CONFIG_PATH = os.getenv(
    "DEBUTADE_CONFIG",
    os.path.abspath(os.path.join(SCRIPT_DIR, "..", "config.json")),
)


def load_config(config_path, section_key="contributie"):
    """Laad de configuratie uit config.json."""
    if not os.path.exists(config_path):
        raise FileNotFoundError(f"Configuratiebestand niet gevonden: {config_path}")

    with open(config_path, "r", encoding="utf-8") as config_file:
        root_config = json.load(config_file)

    if section_key not in root_config:
        raise KeyError(f"Configuratiesectie ontbreekt: {section_key}")

    config = root_config[section_key]
    shared = root_config.get("shared", {})
    for key in ("backup_directory", "log_directory", "log_level"):
        if key in shared:
            config[key] = shared[key]
    if shared.get("bank_excel_file_name") and not config.get("bank_excel_file_name"):
        config["bank_excel_file_name"] = shared["bank_excel_file_name"]

    if shared.get("grootboek_directory") and config.get("bank_excel_file_name"):
        config["bank_excel_file_path"] = os.path.join(
            shared["grootboek_directory"],
            config["bank_excel_file_name"],
        )

    required_keys = [
        "ledenbestand_path",
        "leden_sheet_personen",
        "leden_sheet_betaald",
        "bank_excel_file_name",
        "bank_sheet_name",
        "tags",
    ]
    for key in required_keys:
        if key not in config:
            raise KeyError(f"Configuratiesleutel ontbreekt: {key}")

    return config


try:
    config = load_config(CONFIG_PATH)
except (FileNotFoundError, KeyError) as exc:
    print(f"WAARSCHUWING: {exc}")
    config = {
        "ledenbestand_path": r"C:\pad\naar\Ledenbestand.xlsx",
        "leden_sheet_personen": "personen",
        "leden_sheet_betaald": "betaald",
        "bank_excel_file_name": "Debutade boekjaar bank 2026.xlsx",
        "bank_sheet_name": "Bankrekening",
        "tags": ["contributie-volwassenen", "contributie-jeugd"],
        "backup_directory": os.path.join(SCRIPT_DIR, "backup"),
        "log_directory": os.path.join(SCRIPT_DIR, "logs"),
        "log_level": "INFO",
    }

LEDENBESTAND_PATH = config["ledenbestand_path"]
LEDEN_SHEET_PERSONEN = config["leden_sheet_personen"]
LEDEN_SHEET_BETAALD = config["leden_sheet_betaald"]
BANK_EXCEL_PATH = config.get("bank_excel_file_path") or config.get("bank_excel_file_name")
BANK_SHEET_NAME = config["bank_sheet_name"]
TAGS = [str(tag).strip().lower() for tag in config.get("tags", [])]
TAG_TARGETS = {
    str(key).strip().lower(): float(value)
    for key, value in config.get(
        "tag_targets",
        {
            "8000": 290.0,
            "8001": 185.0,
        },
    ).items()
}
TAG_TARGET_ORDER = list(TAG_TARGETS.keys())
BACKUP_DIRECTORY = config.get("backup_directory", os.path.join(SCRIPT_DIR, "backup"))
LOG_DIRECTORY = config.get("log_directory", os.path.join(SCRIPT_DIR, "logs"))
LOG_LEVEL = config.get("log_level", "INFO")
MAIN_APP_URL = os.getenv("MAIN_APP_URL", "").strip()

if not os.path.exists(LOG_DIRECTORY):
    os.makedirs(LOG_DIRECTORY)

log_file = os.path.join(LOG_DIRECTORY, f"contributie_{datetime.now().strftime('%Y%m%d')}.log")
logging.basicConfig(
    level=getattr(logging, LOG_LEVEL),
    format="%(asctime)s - %(levelname)s - %(message)s",
    handlers=[
        logging.FileHandler(log_file, encoding="utf-8"),
        logging.StreamHandler(),
    ],
)


def normalize_header(value):
    return str(value or "").strip().lower().replace(" ", "").replace("-", "")


def normalize_account(value):
    if value is None:
        return ""
    if isinstance(value, (int, float)):
        return str(int(value))
    return str(value).strip().replace(" ", "")


def parse_amount(value):
    if value is None:
        return 0.0
    if isinstance(value, (int, float)):
        return float(value)

    text = str(value).strip()
    if not text:
        return 0.0

    # Verwijder valutatekens en niet-numerieke tekens behalve . , -
    cleaned = re.sub(r"[^0-9,\.\-]", "", text)
    if not cleaned:
        return 0.0

    # Als er een komma en punt zijn, ga uit van punt als duizendtallen
    if "," in cleaned and "." in cleaned:
        cleaned = cleaned.replace(".", "")

    # Gebruik punt als decimaal scheiding
    cleaned = cleaned.replace(",", ".")
    try:
        return float(cleaned)
    except ValueError:
        return 0.0


def extract_tag_code(value):
    text = str(value or "").strip().lower()
    if not text:
        return ""
    if ";" in text:
        text = text.split(";", 1)[0]
    return text.strip()


def build_header_map(sheet):
    headers = {}
    for idx, cell in enumerate(sheet[1]):
        normalized = normalize_header(cell.value)
        if normalized:
            headers[normalized] = idx
    return headers


def get_cell_value(row, header_map, *names):
    for name in names:
        idx = header_map.get(normalize_header(name))
        if idx is not None and idx < len(row):
            return row[idx]
    return None


def read_sheet_rows(file_path, sheet_name):
    if not os.path.exists(file_path):
        return [], f"Excel bestand niet gevonden: {file_path}"

    wb = None
    temp_path = None
    try:
        wb = load_workbook(file_path, read_only=True, data_only=True)
    except PermissionError:
        # Bestand is waarschijnlijk open in Excel of door OneDrive gelocked.
        try:
            temp_dir = os.path.join(SCRIPT_DIR, "cache")
            os.makedirs(temp_dir, exist_ok=True)
            temp_file = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx", dir=temp_dir)
            temp_file.close()
            temp_path = temp_file.name
            shutil.copy2(file_path, temp_path)
            wb = load_workbook(temp_path, read_only=True, data_only=True)
        except Exception:
            return [], f"Bestand is in gebruik: {file_path}. Sluit Excel en probeer opnieuw."
    except Exception as exc:  # noqa: BLE001
        return [], f"Kan Excel bestand niet openen: {file_path} ({exc})"

    if sheet_name not in wb.sheetnames:
        wb.close()
        if temp_path and os.path.exists(temp_path):
            os.remove(temp_path)
        return [], f"Tabblad niet gevonden: {sheet_name}"

    sheet = wb[sheet_name]
    header_map = build_header_map(sheet)
    rows = list(sheet.iter_rows(min_row=2, values_only=True))
    wb.close()
    if temp_path and os.path.exists(temp_path):
        os.remove(temp_path)
    return (rows, header_map), None


def load_ledenbestand():
    errors = []
    persons_map = {}
    betaald_rows = []

    result, error = read_sheet_rows(LEDENBESTAND_PATH, LEDEN_SHEET_PERSONEN)
    if error:
        errors.append(error)
    else:
        rows, header_map = result
        for row in rows:
            member_id = get_cell_value(row, header_map, "ID-lid", "ID lid", "ID")
            achternaam = get_cell_value(row, header_map, "Achternaam")
            contributie = get_cell_value(row, header_map, "Contributie")
            if member_id is None and not achternaam:
                continue
            member_id = str(member_id).strip() if member_id is not None else ""
            persons_map[member_id] = {
                "achternaam": str(achternaam or "").strip(),
                "contributie": str(contributie or "").strip(),
            }

    result, error = read_sheet_rows(LEDENBESTAND_PATH, LEDEN_SHEET_BETAALD)
    if error:
        errors.append(error)
    else:
        rows, header_map = result
        for row in rows:
            member_id = get_cell_value(row, header_map, "ID")
            rekening = get_cell_value(row, header_map, "Rekening", "Rekeningnummer", "Rek.")
            member_id = str(member_id).strip() if member_id is not None else ""
            rekening = normalize_account(rekening)
            if not member_id and not rekening:
                continue
            betaald_rows.append({
                "member_id": member_id,
                "rekening": rekening,
            })

    return persons_map, betaald_rows, errors


def load_contributie_totals():
    if not BANK_EXCEL_PATH or not os.path.exists(BANK_EXCEL_PATH):
        return {}, [f"Bankrekening Excel bestand niet gevonden: {BANK_EXCEL_PATH}"]

    try:
        wb = load_workbook(BANK_EXCEL_PATH, read_only=True, data_only=True)
    except PermissionError:
        return {}, [
            f"Bestand is in gebruik: {BANK_EXCEL_PATH}. Sluit Excel en probeer opnieuw."
        ]
    except Exception as exc:  # noqa: BLE001
        return {}, [f"Kan Excel bestand niet openen: {BANK_EXCEL_PATH} ({exc})"]
    if BANK_SHEET_NAME not in wb.sheetnames:
        wb.close()
        return {}, [f"Bankrekening tabblad niet gevonden: {BANK_SHEET_NAME}"]

    sheet = wb[BANK_SHEET_NAME]
    header_map = build_header_map(sheet)
    totals = {}

    for row in sheet.iter_rows(min_row=2, values_only=True):
        tag_value = get_cell_value(row, header_map, "Tag")
        tag_code = extract_tag_code(tag_value)
        if tag_code not in TAGS:
            continue

        tegenrekening = get_cell_value(row, header_map, "Tegenrekening", "Tegen rekening")
        bedrag = get_cell_value(row, header_map, "Bedrag (EUR)", "Bedrag", "BedragEUR")
        af_bij = get_cell_value(row, header_map, "Af Bij", "Af/Bij")

        rekening_key = normalize_account(tegenrekening)
        if not rekening_key:
            continue

        amount_value = parse_amount(bedrag)

        if str(af_bij or "").strip().lower() == "af":
            amount_value = -amount_value

        if rekening_key not in totals:
            totals[rekening_key] = {}
        totals[rekening_key][tag_code] = totals[rekening_key].get(tag_code, 0.0) + amount_value

    wb.close()
    return totals, []


def build_overview():
    persons_map, betaald_rows, errors = load_ledenbestand()
    totals, bank_errors = load_contributie_totals()
    errors.extend(bank_errors)

    rekening_by_id = {}
    for row in betaald_rows:
        member_id = row.get("member_id", "")
        rekening = row.get("rekening", "")
        if member_id and rekening and member_id not in rekening_by_id:
            rekening_by_id[member_id] = rekening

    records = []
    summary = {
        "volwassenen": {"received": 0.0, "max": 0.0, "count": 0},
        "jeugd": {"received": 0.0, "max": 0.0, "count": 0},
    }
    for member_id, person_info in persons_map.items():
        achternaam = person_info.get("achternaam", "")
        contributie = person_info.get("contributie", "")
        contributie_flag = str(contributie or "").strip().lower()
        if contributie_flag.startswith("j"):
            selected_tags = ["8001"]
            summary_key = "jeugd"
        elif contributie_flag.startswith("v"):
            selected_tags = ["8000"]
            summary_key = "volwassenen"
        else:
            selected_tags = TAG_TARGET_ORDER
            summary_key = None
        
        if summary_key:
            summary[summary_key]["count"] += 1

        rekening = rekening_by_id.get(member_id, "")
        paid_by_tag = totals.get(rekening, {})
        paid_total = sum(paid_by_tag.get(tag, 0.0) for tag in selected_tags)
        bar_items = []
        for tag_code in selected_tags:
            target = TAG_TARGETS.get(tag_code, 0.0)
            paid = paid_by_tag.get(tag_code, 0.0)
            percent = (paid / target * 100) if target else 0.0
            bar_items.append({
                "tag": tag_code,
                "target": target,
                "paid": paid,
                "percent": percent,
                "bar_width": min(100.0, percent),
                "overflow": percent > 100.0,
            })
            if summary_key:
                summary[summary_key]["received"] += paid
                summary[summary_key]["max"] += target
        records.append({
            "member_id": member_id,
            "achternaam": achternaam,
            "rekening": rekening,
            "paid": paid_total,
            "bar_items": bar_items,
        })

    records.sort(key=lambda item: (item.get("achternaam", ""), item.get("member_id", "")))

    with_payment = sum(1 for item in records if item.get("paid", 0) > 0)
    without_payment = len(records) - with_payment

    stats = {
        "total": len(records),
        "with_payment": with_payment,
        "without_payment": without_payment,
    }

    return records, stats, errors, summary


@app.route("/")
def index():
    records, stats, errors, summary = build_overview()
    current_date = datetime.now().strftime("%d-%m-%Y")
    current_user = os.getlogin()
    return render_template(
        "contributie.html",
        records=records,
        stats=stats,
        errors=errors,
        summary=summary,
        current_date=current_date,
        current_user=current_user,
        main_app_url=MAIN_APP_URL,
    )


@app.route("/quit", methods=["POST"])
def quit_app():
    """Sluit de applicatie af."""
    try:
        logging.info("Applicatie wordt afgesloten...")
        logging.info("=" * 70)

        response = jsonify({"success": True, "message": "Applicatie sluit af"})

        def shutdown_server():
            import time
            time.sleep(1)
            logging.info("Flask server wordt beeindigd...")
            os._exit(0)

        import threading
        shutdown_thread = threading.Thread(target=shutdown_server, daemon=True)
        shutdown_thread.start()

        return response, 200
    except Exception as exc:
        logging.error(f"Fout bij afsluiten applicatie: {str(exc)}")
        return jsonify({"success": False, "message": f"Fout: {str(exc)}"}), 500


if __name__ == "__main__":
    app.run(debug=False, host="127.0.0.1", port=5004)
