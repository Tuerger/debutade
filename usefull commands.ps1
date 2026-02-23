schtasks /End /TN "\Debutade\Startup"
schtasks /Run /TN "\Debutade\Startup"
schtasks /Query /TN "\Debutade\Startup" /V /FO LIST

#python processen zoeken en stoppen inclusief de parent
Stop-Process -Id 5848 -Force
Stop-Process -Id 5804 -Force
Net-NetTCPConnection -LocalPort 5003 -State Listen


Get-CimInstance Win32_Process -Filter "ProcessId=5848" | Select ProcessId,ParentProcessId,CommandLine
sc qc DebutadeService
Get-ScheduledTask | Where-Object { $_.TaskName -match 'Debutade|watchdog' -or $_.TaskPath -match 'Debutade' }



Git commando's:
cd C:\Debutade

2) Check of dit al een git-repo is (en welke remote actief is)
git status
git remote -v

3) Stage + commit je lokale wijzigingen
git add -A
git commit -m "Lokale wijzigingen Debutade"