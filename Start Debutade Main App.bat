@echo off
setlocal

rem --- always start in the folder of the batch file ---
cd /d "%~dp0"
set "APP_DIR=%CD%\"


set "APP_DIR=%CD%\"
set "PYTHONW_EXE=%APP_DIR%.venv-main\Scripts\pythonw.exe"
set "PYTHON_EXE=%APP_DIR%.venv-main\Scripts\python.exe"
set "LOG_FILE=%APP_DIR%start-debutade-startup.log"
set "MAIN_APP_HOST=127.0.0.1"

rem --- select python executable ---
if exist "%PYTHONW_EXE%" (
    set "RUN_EXE=%PYTHONW_EXE%"
) else if exist "%PYTHON_EXE%" (
    set "RUN_EXE=%PYTHON_EXE%"
) else (
    set "RUN_EXE=python"
)

echo [%date% %time%] Start Debutade Main App >> "%LOG_FILE%"

rem --- run without START (fails in Scheduled Task) ---
"%RUN_EXE%" "%APP_DIR%app.py" >> "%LOG_FILE%" 2>&1

endlocal
