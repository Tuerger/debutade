"""
Rapporten Debutade - Web Applicatie
===================================

Dashboard met gecombineerde rapportage van kasboek en bankrekening.

Versie: 1.0
Datum: 2026-02-21
"""

from __future__ import annotations

from datetime import datetime, date
import getpass
import json
import logging
import os
import sys
import time
from typing import Any

from flask import Flask, jsonify, redirect, render_template, request, g
from openpyxl import load_workbook


if sys.platform == "win32":
    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except AttributeError:
        import io

        sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")


SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
CONFIG_PATH = os.getenv(
    "DEBUTADE_CONFIG",
    os.path.abspath(os.path.join(SCRIPT_DIR, "..", "config.json")),
)


def load_config(config_path: str) -> dict[str, Any]:
    if not os.path.exists(config_path):
        raise FileNotFoundError(f"Configuratiebestand niet gevonden: {config_path}")

    with open(config_path, "r", encoding="utf-8") as config_file:
        root_config = json.load(config_file)

    shared = root_config.get("shared", {})
    rapporten = root_config.get("rapporten", {})
    bank = root_config.get("bankrekening", {})
    kas = root_config.get("kasboek", {})

    grootboek_directory = shared.get("grootboek_directory", "")

    bank_excel_file_name = (
        rapporten.get("bank_excel_file_name")
        or bank.get("excel_file_name")
        or shared.get("bank_excel_file_name")
        or ""
    )
    kas_excel_file_name = (
        rapporten.get("kas_excel_file_name")
        or kas.get("excel_file_name")
        or ""
    )

    bank_sheets = rapporten.get("bank_sheets") or bank.get("required_sheets") or [
        "Bankrekening",
        "Spaarrekening 1",
        "Spaarrekening 2",
    ]
    kas_sheet_name = rapporten.get("kas_sheet_name") or kas.get("excel_sheet_name") or "Kas"

    def build_path(file_name: str) -> str:
        if not file_name:
            return ""
        if os.path.isabs(file_name):
            return file_name
        if grootboek_directory:
            return os.path.join(grootboek_directory, file_name)
        return file_name

    return {
        "bank_excel_path": build_path(bank_excel_file_name),
        "kas_excel_path": build_path(kas_excel_file_name),
        "bank_sheets": bank_sheets,
        "kas_sheet_name": kas_sheet_name,
        "log_directory": shared.get("log_directory", os.path.join(SCRIPT_DIR, "logs")),
        "log_level": shared.get("log_level", "INFO"),
        "main_app_url": os.getenv("MAIN_APP_URL", "").strip(),
    }


try:
    config = load_config(CONFIG_PATH)
except (FileNotFoundError, KeyError, json.JSONDecodeError) as exc:
    print(f"WAARSCHUWING: {exc}")
    config = {
        "bank_excel_path": "",
        "kas_excel_path": "",
        "bank_sheets": ["Bankrekening", "Spaarrekening 1", "Spaarrekening 2"],
        "kas_sheet_name": "Kas",
        "log_directory": os.path.join(SCRIPT_DIR, "logs"),
        "log_level": "INFO",
        "main_app_url": os.getenv("MAIN_APP_URL", "").strip(),
    }


LOG_DIRECTORY = config["log_directory"]
LOG_LEVEL = config["log_level"]
MAIN_APP_URL = config["main_app_url"]

app = Flask(
    __name__,
    template_folder=os.path.join(SCRIPT_DIR, "templates"),
    static_folder=os.path.join(SCRIPT_DIR, "static"),
)
app.config["TEMPLATES_AUTO_RELOAD"] = True

CACHE_TTL_SECONDS = int(os.getenv("DEBUTADE_CACHE_TTL_SECONDS", "20"))
BENCHMARK_ENABLED = os.getenv("DEBUTADE_BENCHMARK", "1") == "1"
RUNTIME_CACHE: dict[tuple[Any, ...], dict[str, Any]] = {}


def _file_signature(file_path: str) -> tuple[int, int] | None:
    if not file_path or not os.path.exists(file_path):
        return None
    stat = os.stat(file_path)
    return (stat.st_mtime_ns, stat.st_size)


def _cache_get(key: tuple[Any, ...]) -> Any:
    entry = RUNTIME_CACHE.get(key)
    if not entry:
        return None

    if time.time() - entry["ts"] > CACHE_TTL_SECONDS:
        RUNTIME_CACHE.pop(key, None)
        return None

    return entry["value"]


def _cache_set(key: tuple[Any, ...], value: Any) -> Any:
    RUNTIME_CACHE[key] = {"ts": time.time(), "value": value}
    return value


def normalize_text(value: Any) -> str:
    return str(value or "").strip()


def parse_date(value: Any) -> datetime | None:
    if isinstance(value, datetime):
        return value
    if isinstance(value, date):
        return datetime(value.year, value.month, value.day)
    if value is None:
        return None

    raw = normalize_text(value)
    if not raw:
        return None

    for fmt in ("%d-%m-%Y", "%Y-%m-%d", "%d/%m/%Y", "%d-%m-%y", "%Y/%m/%d"):
        try:
            return datetime.strptime(raw, fmt)
        except ValueError:
            continue
    return None


def parse_amount(value: Any) -> float | None:
    if isinstance(value, (int, float)):
        return float(value)

    raw = normalize_text(value)
    if not raw:
        return None

    raw = raw.replace("€", "").replace(" ", "")
    if raw.count(",") == 1 and raw.count(".") > 1:
        raw = raw.replace(".", "")
    raw = raw.replace(".", "").replace(",", ".")

    try:
        return float(raw)
    except ValueError:
        return None


def source_from_sheet(sheet_name: str, is_kas: bool = False) -> str:
    if is_kas:
        return "Kas"

    lower_sheet = normalize_text(sheet_name).lower()
    if "spaarrekening 1" in lower_sheet:
        return "Spaarrekening 1"
    if "spaarrekening 2" in lower_sheet:
        return "Spaarrekening 2"
    if "bank" in lower_sheet:
        return "Bankrekening"
    return normalize_text(sheet_name) or "Onbekend"


def read_transactions_from_sheet(file_path: str, sheet_name: str, is_kas: bool = False) -> list[dict[str, Any]]:
    records: list[dict[str, Any]] = []
    workbook = None

    if not file_path or not os.path.exists(file_path):
        return records

    try:
        workbook = load_workbook(file_path, read_only=True, data_only=True)
        if sheet_name not in workbook.sheetnames:
            return records

        sheet = workbook[sheet_name]
        source = source_from_sheet(sheet_name, is_kas=is_kas)

        header_row = next(sheet.iter_rows(min_row=1, max_row=1, values_only=True), None)
        header_map: dict[str, int] = {}
        if header_row:
            for index, cell_value in enumerate(header_row):
                key = normalize_text(cell_value).lower()
                if key and key not in header_map:
                    header_map[key] = index

        def get_value(row: tuple[Any, ...], header_names: tuple[str, ...], fallback_index: int | None = None) -> Any:
            for header_name in header_names:
                idx = header_map.get(header_name.lower())
                if idx is not None and idx < len(row):
                    return row[idx]
            if fallback_index is not None and fallback_index < len(row):
                return row[fallback_index]
            return None

        for row in sheet.iter_rows(min_row=2, values_only=True):
            if not row:
                continue

            txn_date = parse_date(get_value(row, ("datum",), 0))
            af_bij = normalize_text(get_value(row, ("af bij", "af/bij"), 5))
            amount = parse_amount(get_value(row, ("bedrag (eur)", "bedrag", "amount"), 6))

            if amount is None or af_bij not in {"Af", "Bij"}:
                continue

            sign = -1.0 if af_bij == "Af" else 1.0
            signed_amount = round(sign * amount, 2)
            month_key = txn_date.strftime("%Y-%m") if txn_date else "Onbekend"

            records.append(
                {
                    "date": txn_date.strftime("%Y-%m-%d") if txn_date else "",
                    "month": month_key,
                    "description": normalize_text(get_value(row, ("naam / omschrijving", "omschrijving", "naam"), 1)),
                    "source": source,
                    "af_bij": af_bij,
                    "amount": round(float(amount), 2),
                    "signed_amount": signed_amount,
                    "bon": normalize_text(get_value(row, ("bon",), 10)),
                    "tag": normalize_text(get_value(row, ("tag",), 11)) or "(Geen tag)",
                    "mededelingen": normalize_text(get_value(row, ("mededelingen",), 8)),
                }
            )

        return records
    except Exception as exc:
        logging.error("Fout bij lezen sheet %s uit %s: %s", sheet_name, file_path, exc)
        return records
    finally:
        if workbook:
            workbook.close()


def load_all_transactions() -> tuple[list[dict[str, Any]], list[str]]:
    bank_file = config.get("bank_excel_path", "")
    kas_file = config.get("kas_excel_path", "")
    bank_sheets = tuple(config.get("bank_sheets", []))
    kas_sheet_name = config.get("kas_sheet_name", "Kas")

    cache_key = (
        "all_transactions",
        bank_file,
        _file_signature(bank_file),
        kas_file,
        _file_signature(kas_file),
        bank_sheets,
        kas_sheet_name,
    )
    cached = _cache_get(cache_key)
    if cached is not None:
        return cached

    all_records: list[dict[str, Any]] = []
    warnings: list[str] = []

    if not bank_file or not os.path.exists(bank_file):
        warnings.append("Bank Excel bestand niet gevonden of niet ingesteld.")
    if not kas_file or not os.path.exists(kas_file):
        warnings.append("Kas Excel bestand niet gevonden of niet ingesteld.")

    for sheet_name in bank_sheets:
        all_records.extend(read_transactions_from_sheet(bank_file, sheet_name, is_kas=False))

    all_records.extend(read_transactions_from_sheet(kas_file, kas_sheet_name, is_kas=True))

    all_records.sort(key=lambda item: (item["date"], item["description"]), reverse=True)
    return _cache_set(cache_key, (all_records, warnings))


def get_report_payload_cached() -> dict[str, Any]:
    records, warnings = load_all_transactions()
    bank_file = config.get("bank_excel_path", "")
    kas_file = config.get("kas_excel_path", "")

    cache_key = (
        "report_payload",
        bank_file,
        _file_signature(bank_file),
        kas_file,
        _file_signature(kas_file),
    )
    cached = _cache_get(cache_key)
    if cached is not None:
        return cached

    months = sorted({row["month"] for row in records if row["month"] and row["month"] != "Onbekend"})
    tags = sorted({row["tag"] for row in records if row["tag"]})
    sources = sorted({row["source"] for row in records if row["source"]})

    payload = {
        "transactions": records,
        "filters": {
            "months": months,
            "tags": tags,
            "sources": sources,
        },
        "warnings": warnings,
    }
    return _cache_set(cache_key, payload)


@app.before_request
def log_request() -> None:
    logging.info("REQUEST %s %s %s", request.remote_addr, request.method, request.path)
    if BENCHMARK_ENABLED:
        g._start_time = time.perf_counter()


@app.after_request
def benchmark_request(response):
    if BENCHMARK_ENABLED and hasattr(g, "_start_time"):
        elapsed_ms = (time.perf_counter() - g._start_time) * 1000
        logging.info("PERF %s %s %s %.1fms", request.method, request.path, response.status_code, elapsed_ms)
    return response


@app.route("/")
def index():
    return render_template(
        "index.html",
        current_date=datetime.now().strftime("%d-%m-%Y"),
        current_user=getpass.getuser(),
        main_app_url=MAIN_APP_URL,
    )


@app.route("/api/report-data")
def report_data():
    return jsonify(get_report_payload_cached())


@app.route("/quit", methods=["POST"])
def quit_application():
    try:
        user = getpass.getuser()
        logging.info("APPLICATIE AFGESLOTEN | Gebruiker: %s", user)
        logging.info("=" * 70)

        response = jsonify({"success": True, "message": "Applicatie sluit af"})

        def shutdown_server() -> None:
            import time

            time.sleep(1)
            logging.info("Flask server wordt beeindigd...")
            os._exit(0)

        import threading

        shutdown_thread = threading.Thread(target=shutdown_server, daemon=True)
        shutdown_thread.start()

        return response, 200
    except Exception as exc:
        logging.error("Fout bij afsluiten applicatie: %s", str(exc))
        return jsonify({"success": False, "message": f"Fout: {str(exc)}"}), 500


@app.route("/settings")
def settings():
    if MAIN_APP_URL:
        return redirect(f"{MAIN_APP_URL}/settings")
    return jsonify({"success": False, "message": "Instellingen zijn alleen beschikbaar via de hoofdapp."}), 403


if __name__ == "__main__":
    if not os.path.exists(LOG_DIRECTORY):
        try:
            os.makedirs(LOG_DIRECTORY)
        except Exception as exc:
            print(f"FOUT: Kan log directory niet aanmaken: {LOG_DIRECTORY}")
            print(f"Details: {str(exc)}")
            exit(1)

    log_file_path = os.path.join(LOG_DIRECTORY, "rapporten_webapp_log.txt")
    logging.basicConfig(
        filename=log_file_path,
        level=getattr(logging, str(LOG_LEVEL).upper(), logging.INFO),
        format="%(asctime)s - %(levelname)s - %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
    )

    logging.info("=" * 70)
    logging.info("RAPPORTEN START")
    logging.info("Bank bestand: %s", config.get("bank_excel_path", ""))
    logging.info("Kas bestand: %s", config.get("kas_excel_path", ""))
    logging.info("=" * 70)

    port = int(os.getenv("DEBUTADE_APP_PORT", "5004"))
    app.run(debug=False, host="127.0.0.1", port=port, use_reloader=False)
