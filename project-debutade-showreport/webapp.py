"""
Showreport Debutade - Web Applicatie
===================================

Toont een Power BI rapport in een embedded frame met Debutade layout.

Versie: 1.0
Datum: 2026-02-09
Auteur: Eric G.
"""

from flask import Flask, render_template, jsonify, redirect, request
from datetime import datetime
import os
import json
import logging
import getpass
import sys

# Fix encoding voor Windows console
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


def load_config(config_path, section_key="showreport"):
    """Laad de configuratie uit een JSON bestand"""
    if not os.path.exists(config_path):
        raise FileNotFoundError(f"Configuratiebestand niet gevonden: {config_path}")

    with open(config_path, "r", encoding="utf-8") as config_file:
        root_config = json.load(config_file)

    config = root_config.get(section_key, {})
    shared = root_config.get("shared", {})
    for key in ("log_directory", "log_level"):
        if key in shared:
            config[key] = shared[key]

    return config


try:
    config = load_config(CONFIG_PATH)
except (FileNotFoundError, KeyError) as e:
    print(f"WAARSCHUWING: {e}")
    config = {}

REPORT_URL = (config.get("report_url") or "").strip()
REPORT_TITLE = (config.get("report_title") or "Debutade Rapport").strip()
LOG_DIRECTORY = config.get("log_directory", os.path.join(SCRIPT_DIR, "logs"))
LOG_LEVEL = config.get("log_level", "INFO")
MAIN_APP_URL = os.getenv("MAIN_APP_URL", "").strip()

app = Flask(
    __name__,
    template_folder=os.path.join(SCRIPT_DIR, "templates"),
    static_folder=os.path.join(SCRIPT_DIR, "..", "static"),
)
app.config["TEMPLATES_AUTO_RELOAD"] = True


@app.before_request
def log_request():
    logging.info("REQUEST %s %s %s", request.remote_addr, request.method, request.path)


@app.route("/")
def index():
    current_date_display = datetime.now().strftime("%d-%m-%Y")
    current_user = getpass.getuser()
    logging.info("Showreport geopend door %s", current_user)
    return render_template(
        "index.html",
        report_url=REPORT_URL,
        report_title=REPORT_TITLE,
        current_date=current_date_display,
        current_user=current_user,
        main_app_url=MAIN_APP_URL,
    )


@app.route("/quit", methods=["POST"])
def quit_application():
    """Stop de applicatie en log dit"""
    try:
        user = getpass.getuser()
        logging.info("APPLICATIE AFGESLOTEN | Gebruiker: %s", user)
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
    except Exception as e:
        logging.error(f"Fout bij afsluiten applicatie: {str(e)}")
        return jsonify({"success": False, "message": f"Fout: {str(e)}"}), 500


@app.route("/settings")
def settings():
    if MAIN_APP_URL:
        return redirect(f"{MAIN_APP_URL}/settings")
    return jsonify({"success": False, "message": "Instellingen zijn alleen beschikbaar via de hoofdapp."}), 403


if __name__ == "__main__":
    if not os.path.exists(LOG_DIRECTORY):
        try:
            os.makedirs(LOG_DIRECTORY)
        except Exception as e:
            print(f"FOUT: Kan log directory niet aanmaken: {LOG_DIRECTORY}")
            print(f"Details: {str(e)}")
            exit(1)

    log_file_path = os.path.join(LOG_DIRECTORY, "showreport_webapp_log.txt")
    logging.basicConfig(
        filename=log_file_path,
        level=getattr(logging, LOG_LEVEL.upper(), logging.INFO),
        format="%(asctime)s - %(levelname)s - %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
    )

    logging.info("=" * 70)
    logging.info("SHOWREPORT START")
    logging.info("Report titel: %s", REPORT_TITLE)
    logging.info("Report URL ingesteld: %s", "ja" if REPORT_URL else "nee")
    logging.info("Log directory: %s", LOG_DIRECTORY)
    logging.info("=" * 70)

    port = int(os.getenv("DEBUTADE_APP_PORT", "5004"))
    app.run(debug=False, host="127.0.0.1", port=port, use_reloader=False)
