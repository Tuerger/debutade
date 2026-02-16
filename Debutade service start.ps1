# --- Debutade taak aanmaken in eigen map (\Debutade\Startup) ---

$taskName   = "\Debutade\Startup"
$pythonExe  = "C:\Debutade\.venv-main\Scripts\python.exe"
$scriptFile = "C:\Debutade\watchdog.py"

# 1) Verwijder bestaande taak als die er is (stil als die niet bestaat)
schtasks /Delete /TN "$taskName" /F 2>$null

# 2) Maak nieuwe taak onder SYSTEM, ONSTART met 30s delay
#    Let op: /TR zonder quotes, omdat er GEEN spaties in de paden zitten
schtasks /Create `
  /TN "$taskName" `
  /SC ONSTART `
  /DELAY 0000:30 `
  /RL HIGHEST `
  /RU SYSTEM `
  /TR $("$pythonExe $scriptFile") `
  /F

# 3) Start direct en toon status
schtasks /Run /TN "$taskName"
schtasks /Query /TN "$taskName" /V /FO LIST