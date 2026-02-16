# Debutade Apps Launcher

A modern graphical application launcher for your Debutade applications.

## Features

- Clean, modern UI with dark theme
- Launch Bankrekening application
- Launch Kasboek application
- **Start/Stop Webservices** - Toggle both webservices (Kasboek & Bankrekening) with a single button
- Hidden console windows for webservices (runs in background)
- Configurable logging system
- Error handling for missing files
- Professional appearance with organized layout
- Custom taskbar icon

## Installation

1. Make sure you have Python 3.7+ installed
2. Install dependencies:
   ```
   pip install -r requirements.txt
   ```

## Usage

### Option 1: Using the executable (Recommended)
```
dist\Debutade Start Apps.exe
```

### Option 2: Direct Python execution
```
python app.py
```

### Option 3: Using batch file
```
Run.bat
```

## Build

You can generate the Windows executable with PyInstaller.

### Prerequisites
- Python 3.7+ (tested with 3.13)
- Dependencies installed: `pip install -r requirements.txt`
- PyInstaller installed: `pip install -U pyinstaller`

### Build using .spec (preferred)
```powershell
cd C:\project-start-debutade-apps
pyinstaller "Debutade Apps Starter.spec" --clean
```

### Build command
```powershell
cd C:\project-start-debutade-apps
python -m PyInstaller --noconfirm --onefile --windowed --icon static/icon.ico --name "Debutade Start Apps" --add-data "static;static" app.py
```

Notes:
- The icon is taken from `static/icon.ico`.
- The `static` folder is bundled so images/icons load correctly when running the EXE.
- The generated executable will be placed in `dist/` as `Debutade Start Apps.exe`.

### Clean up old builds (optional)
If you previously produced a typo-named EXE, you can remove it:
```powershell
Remove-Item "dist/Debutade Styart Apps.exe" -Force
```

## Webservices

The application can start and stop both webservices (Kasboek and Bankrekening) simultaneously:

1. Click **"Start Webservices"** to launch both webservers in the background
2. The button changes to **"Stop Webservices"** once running
3. Click **"Stop Webservices"** to gracefully terminate both webservers
4. A confirmation message appears when services are stopped

Webservices run without visible console windows for a cleaner experience.

## Configuration

Edit `start-debutade.config` to customize logging settings:

```ini
[logging]
level = INFO
file = C:\path\to\your\log\file.log
format = %(asctime)s [%(levelname)s] %(message)s
```

Available log levels: DEBUG, INFO, WARNING, ERROR, CRITICAL

## Files

- `app.py` - Main application code
- `requirements.txt` - Python dependencies
- `Run.bat` - Batch file to easily launch the app
- `start-debutade.config` - Configuration file for logging
- `Start Debutade Webservers.bat` - Batch script for launching webservices
- `Debutade Apps Starter.spec` - Primary PyInstaller spec for building the executable
- `DebutadeAppsStarter.spec` - Alternate spec filename (same packaging intent)
- `Debutade Start Apps.spec` - Spec variant with correct product name
- `Debutade Styart Apps.spec` - Legacy/typo spec (not used)
- `dist\Debutade Start Apps.exe` - Compiled executable with custom icon
- `static\icon.ico` - Application icon
- `static\Header Debutade.jpg` - Header image

## Release

Latest updates:
- 2026-01-09: Added documented build via `.spec` (preferred) using "Debutade Apps Starter.spec" and noted Python/PyInstaller versions used.
- 2026-01-09: Updated Files section to list all available `.spec` files and marked the legacy typo-named spec as not used.
- 2026-01-09: General documentation cleanup for build and usage.
- 2026-01-08: **Fixed config loading in EXE** - The `start-debutade.config` file is now properly bundled with the executable and logging configuration is correctly applied when running the EXE. Updated `.spec` file to include config in `datas` and improved config path resolution logic.
- 2026-01-08: Fixed duplicate launcher issue when pressing **Start Webservices** in the EXE by using system `python` for webapps when frozen.
- 2026-01-08: Added a clear Build section with PyInstaller command and asset bundling.
- 2026-01-08: Corrected executable name to **Debutade Start Apps.exe** and documented optional cleanup of the older typo-named file.

Notes:
- Ensure `python` is available on PATH when running the EXE so webservices can launch.
- Configure logging path in `start-debutade.config` as needed.
- The config file is bundled with the EXE but can be overridden by placing a `start-debutade.config` file next to the executable.
