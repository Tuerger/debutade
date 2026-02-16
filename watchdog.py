# C:\Debutade\watchdog.py

import os
import sys
import time
import subprocess
from datetime import datetime

# --------------------------------------------------------------------------------------
# CONFIG
# --------------------------------------------------------------------------------------
ROOT = r"C:\Debutade"
APP = os.path.join(ROOT, "app.py")
LOGDIR = os.path.join(ROOT, "logs")
WATCHDOG_LOG = os.path.join(LOGDIR, "watchdog.log")

# Gebruik python.exe zolang je wilt zien wat er gebeurt.
# Als alles werkt, kun je pythonw.exe gebruiken.
PYTHON = r"C:\Debutade\.venv-main\Scripts\python.exe"

CHECK_INTERVAL_SEC = 5

# --------------------------------------------------------------------------------------
# HELPERS
# --------------------------------------------------------------------------------------
def ts() -> str:
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")

def ensure_dirs():
    os.makedirs(LOGDIR, exist_ok=True)

def write_watchdog(msg: str):
    """Schrijf veilige logregels, crasht nooit."""
    try:
        ensure_dirs()
        with open(WATCHDOG_LOG, "a", encoding="utf-8") as f:
            f.write(f"{ts()} {msg}\n")
    except Exception:
        pass

# --------------------------------------------------------------------------------------
# START APP
# --------------------------------------------------------------------------------------
def start_app():
    """Start de webservice app.py en log wat er gebeurt."""
    if not os.path.isfile(PYTHON):
        write_watchdog(f"[ERROR] Interpreter niet gevonden: {PYTHON}")
        return None

    if not os.path.isfile(APP):
        write_watchdog(f"[ERROR] APP niet gevonden: {APP}")
        return None

    try:
        cmd = [PYTHON, APP]
        write_watchdog(f"[INFO] Start app: {cmd}")

        proc = subprocess.Popen(
            cmd,
            cwd=ROOT,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE
        )

        write_watchdog(f"[INFO] app.py gestart met PID {proc.pid}")
        return proc

    except Exception as e:
        write_watchdog(f"[ERROR] Start app.py faalde: {e}")
        return None

# --------------------------------------------------------------------------------------
# MAIN
# --------------------------------------------------------------------------------------
def main():
    ensure_dirs()
    write_watchdog("=== WATCHDOG STARTED ===")

    proc = start_app()

    # Eenvoudige watchdog-loop (kan later uitgebreid worden)
    while True:
        time.sleep(CHECK_INTERVAL_SEC)
        # Hier kun je healthchecks toevoegen
        # Voor nu loggen we niet elke keer om bestand klein te houden.


if __name__ == "__main__":
    main()