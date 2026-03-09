"""
Contributie Debutade - Web Applicatie
====================================

Web applicatie voor het koppelen van contributiebetalingen aan leden.

Functionaliteiten:
- Leest ledengegevens uit Ledenbestand.xlsx (tab leden)
- Leest banktransacties uit tab bankrekening en zoekt ID-lid in mededelingen
- Maakt een overzicht met Te innen bedrag, Ontvangen bedrag en status

Versie: 1.0
Datum: 2026-02-16
Auteur: Eric G.
"""

from flask import Flask, render_template, jsonify, request, g
from openpyxl import load_workbook
from datetime import datetime
import logging
import os
import json
import sys
import re
import tempfile
import shutil
import time

# Fix encoding voor Windows console
if sys.platform == "win32":
    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except AttributeError:
        import io
        sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")

app = Flask(__name__)
app.config["TEMPLATES_AUTO_RELOAD"] = True

CACHE_TTL_SECONDS = int(os.getenv("DEBUTADE_CACHE_TTL_SECONDS", "20"))
BENCHMARK_ENABLED = os.getenv("DEBUTADE_BENCHMARK", "1") == "1"
RUNTIME_CACHE = {}

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
        "bank_excel_file_name",
        "bank_sheet_name",
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
        "leden_sheet_name": "leden",
        "bank_excel_file_name": "Debutade boekjaar 2026 Bank.xlsx",
        "bank_sheet_name": "bankrekening",
        "backup_directory": os.path.join(SCRIPT_DIR, "backup"),
        "log_directory": os.path.join(SCRIPT_DIR, "logs"),
        "log_level": "INFO",
    }

LEDENBESTAND_PATH = config["ledenbestand_path"]
LEDEN_SHEET_NAME = (
    config.get("leden_sheet_name")
    or config.get("leden_sheet_leden")
    or "leden"
)
BANK_EXCEL_PATH = config.get("bank_excel_file_path") or config.get("bank_excel_file_name")
BANK_SHEET_NAME = config["bank_sheet_name"]
MANUAL_TRANSACTION_MAPPINGS = config.get("manual_transaction_mappings", {})
MANUAL_PAID_OVERRIDES = config.get("manual_paid_overrides", {})
BANK_EXCEL_FALLBACK_BASENAMES = [
    "Debutade boekjaar 2026 Bank",
    "Debutade boekjaar bank 2026",
]
SHARED_BANK_EXCEL_FILE_NAME = config.get("bank_excel_file_name", "")
ALLOWED_CONTRIBUTIE_TAG_CODES = {"8000", "8001"}
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


def _file_signature(file_path):
    if not file_path or not os.path.exists(file_path):
        return None
    stat = os.stat(file_path)
    return (stat.st_mtime_ns, stat.st_size)


def _cache_get(key):
    entry = RUNTIME_CACHE.get(key)
    if not entry:
        return None

    if time.time() - entry["ts"] > CACHE_TTL_SECONDS:
        RUNTIME_CACHE.pop(key, None)
        return None

    return entry["value"]


def _cache_set(key, value):
    RUNTIME_CACHE[key] = {"ts": time.time(), "value": value}
    return value


def invalidate_runtime_cache():
    RUNTIME_CACHE.clear()


@app.before_request
def _benchmark_start():
    if BENCHMARK_ENABLED:
        g._start_time = time.perf_counter()


@app.after_request
def _benchmark_end(response):
    if BENCHMARK_ENABLED and hasattr(g, "_start_time"):
        elapsed_ms = (time.perf_counter() - g._start_time) * 1000
        logging.info("PERF %s %s %s %.1fms", request.method, request.path, response.status_code, elapsed_ms)
    return response


def normalize_account(value):
    if value is None:
        return ""
    if isinstance(value, (int, float)):
        return str(int(value))
    return str(value).strip().replace(" ", "")


def normalize_lastname(value):
    return str(value or "").strip().lower()


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


def normalize_member_id_4digit(value):
    text = str(value or "").strip()
    if not text:
        return ""

    if re.fullmatch(r"\d+(\.0+)?", text):
        try:
            text = str(int(float(text)))
        except ValueError:
            pass

    exact_match = re.search(r"(?<!\d)(\d{4})(?!\d)", text)
    if exact_match:
        return exact_match.group(1)

    digits_only = "".join(re.findall(r"\d", text))
    if not digits_only:
        return ""
    if len(digits_only) >= 4:
        return digits_only[-4:]
    return digits_only.zfill(4)


def extract_4digit_tokens(value):
    text = str(value or "")
    if not text:
        return set()
    return set(re.findall(r"(?<!\d)\d{4}(?!\d)", text))


def find_manual_mapping_for_transaction(mededelingen):
    if not mededelingen:
        return None

    mededelingen_lower = str(mededelingen).lower()
    for member_id, mapping_fragment in MANUAL_TRANSACTION_MAPPINGS.items():
        fragment_lower = str(mapping_fragment or "").lower()
        if fragment_lower and fragment_lower in mededelingen_lower:
            return member_id

    return None


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

    matched_sheet_name = None
    for current_name in wb.sheetnames:
        if current_name.strip().lower() == str(sheet_name).strip().lower():
            matched_sheet_name = current_name
            break

    if not matched_sheet_name:
        wb.close()
        if temp_path and os.path.exists(temp_path):
            os.remove(temp_path)
        return [], f"Tabblad niet gevonden: {sheet_name}"

    sheet = wb[matched_sheet_name]
    header_map = build_header_map(sheet)
    rows = list(sheet.iter_rows(min_row=2, values_only=True))
    wb.close()
    if temp_path and os.path.exists(temp_path):
        os.remove(temp_path)
    return (rows, header_map), None


def load_ledenbestand():
    records = []
    result, error = read_sheet_rows(LEDENBESTAND_PATH, LEDEN_SHEET_NAME)
    if error:
        return [], [error]

    rows, header_map = result
    for row in rows:
        member_id = get_cell_value(row, header_map, "ID-lid", "ID lid", "ID")
        achternaam = get_cell_value(row, header_map, "Achternaam")
        email = get_cell_value(row, header_map, "Email", "E-mail", "Mail")
        te_innen_bedrag = get_cell_value(
            row,
            header_map,
            "bedrag",
            "Te innen bedrag",
            "Contributie",
        )

        member_id = str(member_id).strip() if member_id is not None else ""
        achternaam = str(achternaam or "").strip()
        email = str(email or "").strip()
        due_amount = parse_amount(te_innen_bedrag)

        if not member_id and not achternaam and not email and due_amount == 0:
            continue

        if not member_id:
            continue

        records.append(
            {
                "member_id": member_id,
                "member_id_4digit": normalize_member_id_4digit(member_id),
                "achternaam": achternaam,
                "email": email,
                "due_amount": due_amount,
                "received_amount": 0.0,
                "opmerking": "",
                "status_icon": "❌",
                "status_label": "Nog niets ontvangen",
                "status_class": "status-none",
            }
        )

    return records, []


def load_bank_transactions():
    bank_excel_path = resolve_bank_excel_path()
    result, error = read_sheet_rows(bank_excel_path, BANK_SHEET_NAME)
    if error:
        return [], [error]

    rows, header_map = result
    transactions = []
    for row in rows:
        tag_value = get_cell_value(row, header_map, "Tag", "Tags")
        tag_code = extract_tag_code(tag_value)
        if tag_code not in ALLOWED_CONTRIBUTIE_TAG_CODES:
            continue

        mededelingen = get_cell_value(
            row,
            header_map,
            "Mededelingen",
            "Omschrijving",
            "Beschrijving",
        )
        bedrag = get_cell_value(row, header_map, "Bedrag", "Bedrag (EUR)", "BedragEUR")
        af_bij = get_cell_value(row, header_map, "Af Bij", "Af/Bij")

        amount_value = parse_amount(bedrag)
        if str(af_bij or "").strip().lower() == "af":
            amount_value = -amount_value

        if not mededelingen and amount_value == 0:
            continue

        transactions.append(
            {
                "mededelingen": str(mededelingen or ""),
                "amount": amount_value,
            }
        )

    return transactions, []


def build_transaction_totals_by_member_id_4digit(transactions):
    totals_by_member_id_4digit = {}

    for transaction in transactions:
        amount = transaction.get("amount", 0.0)
        mededelingen = transaction.get("mededelingen", "")
        id_tokens = extract_4digit_tokens(mededelingen)

        for token in id_tokens:
            totals_by_member_id_4digit[token] = totals_by_member_id_4digit.get(token, 0.0) + amount

    return totals_by_member_id_4digit


def build_name_pattern(name_value):
    achternaam = str(name_value or "").strip()
    if not achternaam:
        return None

    if len(achternaam) < 3:
        return None

    escaped = re.escape(achternaam)
    escaped = escaped.replace(r"\ ", r"\s+")
    pattern = rf"(?<![A-Za-zÀ-ÿ]){escaped}(?![A-Za-zÀ-ÿ])"
    return re.compile(pattern, flags=re.IGNORECASE)


def calculate_received_by_name_fallback(achternaam, transactions):
    pattern = build_name_pattern(achternaam)
    if pattern is None:
        return 0.0, False, 0

    fallback_amount = 0.0
    found = False
    matched_count = 0
    for transaction in transactions:
        mededelingen = str(transaction.get("mededelingen", ""))
        if not mededelingen:
            continue

        if pattern.search(mededelingen):
            fallback_amount += transaction.get("amount", 0.0)
            found = True
            matched_count += 1

    return round(fallback_amount, 2), found, matched_count


def resolve_bank_excel_path():
    if BANK_EXCEL_PATH and os.path.exists(BANK_EXCEL_PATH):
        return BANK_EXCEL_PATH

    if SHARED_BANK_EXCEL_FILE_NAME and BANK_EXCEL_PATH:
        bank_dir = os.path.dirname(BANK_EXCEL_PATH)
        shared_candidate = os.path.join(bank_dir, SHARED_BANK_EXCEL_FILE_NAME)
        if os.path.exists(shared_candidate):
            return shared_candidate

    search_dirs = []
    if BANK_EXCEL_PATH:
        bank_dir = os.path.dirname(BANK_EXCEL_PATH)
        if bank_dir:
            search_dirs.append(bank_dir)

    if SCRIPT_DIR not in search_dirs:
        search_dirs.append(SCRIPT_DIR)

    for search_dir in search_dirs:
        for fallback_basename in BANK_EXCEL_FALLBACK_BASENAMES:
            for extension in (".xlsx", ".xlsm", ".xls"):
                candidate = os.path.join(search_dir, f"{fallback_basename}{extension}")
                if os.path.exists(candidate):
                    return candidate

        try:
            for filename in os.listdir(search_dir):
                filename_lower = filename.lower()
                if "debutade" in filename_lower and "boekjaar" in filename_lower and "bank" in filename_lower and "2026" in filename_lower:
                    wildcard_candidate = os.path.join(search_dir, filename)
                    if os.path.isfile(wildcard_candidate):
                        return wildcard_candidate
        except OSError:
            continue

    return BANK_EXCEL_PATH


def build_overview():
    bank_excel_path = resolve_bank_excel_path()
    cache_key = (
        "contributie_overview",
        _file_signature(LEDENBESTAND_PATH),
        _file_signature(bank_excel_path),
        _file_signature(CONFIG_PATH),
    )
    cached = _cache_get(cache_key)
    if cached is not None:
        return cached

    records, errors = load_ledenbestand()
    transactions, bank_errors = load_bank_transactions()
    errors.extend(bank_errors)
    transaction_totals_by_4digit = build_transaction_totals_by_member_id_4digit(transactions)

    transaction_totals_by_manual = {}
    for tx in transactions:
        manual_member_id = find_manual_mapping_for_transaction(tx.get("mededelingen", ""))
        if manual_member_id:
            transaction_totals_by_manual[manual_member_id] = (
                transaction_totals_by_manual.get(manual_member_id, 0.0) + tx.get("amount", 0.0)
            )

    matched_transaction_keys = set()

    exact_count = 0
    partial_count = 0
    none_count = 0
    neutral_count = 0

    for record in records:
        member_id = str(record.get("member_id", "")).strip()
        member_id_4digit = record.get("member_id_4digit", "")
        achternaam = record.get("achternaam", "")
        due_amount = record.get("due_amount", 0.0)
        received_amount = transaction_totals_by_4digit.get(member_id_4digit, 0.0)

        if abs(received_amount) < 0.005:
            received_amount = transaction_totals_by_manual.get(record.get("member_id"), 0.0)

        received_amount = round(received_amount, 2)
        record["received_amount"] = received_amount
        record["opmerking"] = ""
        record["manual_paid_override"] = False

        tx_matched_by_id = set()
        tx_matched_by_fallback = set()

        override_entry = MANUAL_PAID_OVERRIDES.get(member_id)
        if isinstance(override_entry, dict) and override_entry.get("marked_paid"):
            reason = str(override_entry.get("reason", "")).strip()
            record["manual_paid_override"] = True
            if due_amount > 0:
                record["received_amount"] = max(received_amount, due_amount)
            record["opmerking"] = (
                f"Handmatig als betaald gemarkeerd: {reason}"
                if reason
                else "Handmatig als betaald gemarkeerd"
            )
            record["status_icon"] = "✅"
            record["status_label"] = "Handmatig gemarkeerd als betaald"
            record["status_class"] = "status-ok"
            exact_count += 1
            continue

        if not member_id_4digit:
            record["opmerking"] = "ID-lid bevat geen 4 cijfers"
            record["status_icon"] = "❌"
            record["status_label"] = "Geen geldig 4-cijferig ID"
            record["status_class"] = "status-none"
            none_count += 1
            continue

        if abs(due_amount) < 0.005 and abs(received_amount) < 0.005:
            record["opmerking"] = "Geen contributie verschuldigd"
            record["status_icon"] = "🟣"
            record["status_label"] = "Niet van toepassing (0/0)"
            record["status_class"] = "status-neutral"
            neutral_count += 1
            continue

        if abs(received_amount) < 0.005:
            fallback_amount, fallback_found, fallback_count = calculate_received_by_name_fallback(
                achternaam,
                transactions,
            )
            if fallback_found and abs(fallback_amount) >= 0.005:
                received_amount = fallback_amount
                record["received_amount"] = received_amount
                tx_label = "transactie" if fallback_count == 1 else "transacties"
                record["opmerking"] = (
                    f"Geen 4-cijferig ID ({member_id_4digit}) gevonden; "
                    f"backup match op achternaam '{achternaam}' ({fallback_count} {tx_label})"
                )

                for i, tx in enumerate(transactions):
                    pattern = build_name_pattern(achternaam)
                    if pattern and pattern.search(str(tx.get("mededelingen", ""))):
                        tx_matched_by_fallback.add(i)
            else:
                record["opmerking"] = f"Geen 4-cijferig ID ({member_id_4digit}) gevonden in mededelingen"
                record["status_icon"] = "❌"
                record["status_label"] = "Nog niets ontvangen"
                record["status_class"] = "status-none"
                none_count += 1
                continue

        if abs(received_amount) >= 0.005:
            for i, tx in enumerate(transactions):
                if member_id_4digit in extract_4digit_tokens(tx.get("mededelingen", "")):
                    tx_matched_by_id.add(i)

        # Track manual mappings
        tx_matched_by_manual = set()
        if record.get("member_id") in transaction_totals_by_manual and abs(received_amount - transaction_totals_by_manual[record.get("member_id")]) < 0.005:
            for i, tx in enumerate(transactions):
                if find_manual_mapping_for_transaction(tx.get("mededelingen", "")) == record.get("member_id"):
                    tx_matched_by_manual.add(i)
            if tx_matched_by_manual and not record["opmerking"]:
                record["opmerking"] = "Handmatig gematched via config"

        matched_transaction_keys.update(tx_matched_by_id)
        matched_transaction_keys.update(tx_matched_by_fallback)
        matched_transaction_keys.update(tx_matched_by_manual)

        if abs(received_amount - due_amount) <= 0.009:
            record["status_icon"] = "✅"
            record["status_label"] = "Volledig ontvangen"
            record["status_class"] = "status-ok"
            exact_count += 1
            continue

        if received_amount < due_amount:
            record["status_icon"] = "🔵"
            record["status_label"] = "Gedeeltelijk ontvangen"
            record["status_class"] = "status-partial"
            partial_count += 1
            continue

        record["status_icon"] = "✅"
        record["status_label"] = "Ontvangen (meer dan te innen)"
        record["status_class"] = "status-ok"
        exact_count += 1

    unmatched_transactions = []
    for i, tx in enumerate(transactions):
        if i not in matched_transaction_keys:
            unmatched_transactions.append(tx)

    status_sort_order = {
        "status-none": 0,
        "status-partial": 1,
        "status-ok": 2,
        "status-neutral": 3,
    }
    records.sort(
        key=lambda item: (
            status_sort_order.get(item.get("status_class", ""), 99),
            item.get("achternaam", "").lower(),
            item.get("member_id", ""),
        )
    )

    stats = {
        "total": len(records),
        "exact": exact_count,
        "partial": partial_count,
        "none": none_count,
        "neutral": neutral_count,
        "total_due": sum(item.get("due_amount", 0.0) for item in records),
        "total_received": sum(item.get("received_amount", 0.0) for item in records),
        "unmatched_count": len(unmatched_transactions),
        "unmatched_total": round(sum(tx.get("amount", 0.0) for tx in unmatched_transactions), 2),
    }

    return _cache_set(cache_key, (records, stats, errors, unmatched_transactions))


@app.route("/")
def index():
    records, stats, errors, unmatched_transactions = build_overview()
    current_date = datetime.now().strftime("%d-%m-%Y")
    current_user = os.getlogin()
    return render_template(
        "contributie.html",
        records=records,
        stats=stats,
        errors=errors,
        unmatched_transactions=unmatched_transactions,
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


@app.route("/save_manual_mapping", methods=["POST"])
def save_manual_mapping():
    """Slaat een handmatige mapping op van mededelingen fragment naar lid nummer."""
    try:
        data = request.get_json()
        member_id = str(data.get("member_id", "")).strip()
        mededelingen = str(data.get("mededelingen", "")).strip()

        if not member_id or not mededelingen:
            return jsonify({"success": False, "error": "Lid nummer en mededelingen zijn verplicht"}), 400

        # Laad config
        with open(CONFIG_PATH, "r", encoding="utf-8") as f:
            config = json.load(f)

        # Zorg ervoor dat de manual_transaction_mappings sectie bestaat
        if "contributie" not in config:
            config["contributie"] = {}
        if "manual_transaction_mappings" not in config["contributie"]:
            config["contributie"]["manual_transaction_mappings"] = {}

        # Voeg de mapping toe
        config["contributie"]["manual_transaction_mappings"][member_id] = mededelingen

        # Update runtime mapping direct, zodat herstart niet nodig is.
        MANUAL_TRANSACTION_MAPPINGS[member_id] = mededelingen

        # Schrijf config terug naar bestand
        with open(CONFIG_PATH, "w", encoding="utf-8") as f:
            json.dump(config, f, indent=4, ensure_ascii=False)

        invalidate_runtime_cache()

        logging.info(f"Manual mapping opgeslagen: lid {member_id} -> '{mededelingen[:50]}...'")
        return jsonify({"success": True, "message": "Mapping opgeslagen"}), 200

    except Exception as exc:
        logging.error(f"Fout bij opslaan manual mapping: {str(exc)}")
        return jsonify({"success": False, "error": f"Fout: {str(exc)}"}), 500


@app.route("/save_paid_override", methods=["POST"])
def save_paid_override():
    """Sla handmatige betaalstatus op per lid met reden."""
    try:
        data = request.get_json() or {}
        member_id = str(data.get("member_id", "")).strip()
        marked_paid = bool(data.get("marked_paid", False))
        reason = str(data.get("reason", "")).strip()

        if not member_id:
            return jsonify({"success": False, "error": "Lid nummer is verplicht"}), 400

        if marked_paid and not reason:
            return jsonify({"success": False, "error": "Reden is verplicht bij handmatig betaald"}), 400

        with open(CONFIG_PATH, "r", encoding="utf-8") as f:
            config_data = json.load(f)

        if "contributie" not in config_data:
            config_data["contributie"] = {}
        if "manual_paid_overrides" not in config_data["contributie"]:
            config_data["contributie"]["manual_paid_overrides"] = {}

        if marked_paid:
            config_data["contributie"]["manual_paid_overrides"][member_id] = {
                "marked_paid": True,
                "reason": reason,
                "updated_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            }
            MANUAL_PAID_OVERRIDES[member_id] = config_data["contributie"]["manual_paid_overrides"][member_id]
            message = "Lid handmatig als betaald gemarkeerd"
        else:
            config_data["contributie"]["manual_paid_overrides"].pop(member_id, None)
            MANUAL_PAID_OVERRIDES.pop(member_id, None)
            message = "Handmatige betaald-markering verwijderd"

        with open(CONFIG_PATH, "w", encoding="utf-8") as f:
            json.dump(config_data, f, indent=4, ensure_ascii=False)

        invalidate_runtime_cache()
        logging.info("Paid override bijgewerkt voor lid %s | marked_paid=%s", member_id, marked_paid)
        return jsonify({"success": True, "message": message}), 200

    except Exception as exc:
        logging.error("Fout bij opslaan paid override: %s", str(exc))
        return jsonify({"success": False, "error": f"Fout: {str(exc)}"}), 500


if __name__ == "__main__":
    app.run(debug=False, host="127.0.0.1", port=5004)
