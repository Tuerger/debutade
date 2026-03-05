# Transactie Toevoegen - Handleiding

## Hoe te gebruiken

### Stap 1: Start de applicatie
Zorg dat je in de project folder bent en voer uit:
```bash
python webapp.py
```

De applicatie start op `http://127.0.0.1:5005`

### Stap 2: Selecteer de juiste sheet
In het filtermenu selecteer je de sheet waar je transacties wilt toevoegen:
- **Bankrekening** - ING bankrekening
- **Spaarrekening 1** - Spaarrekening 1
- **Spaarrekening 2** - Spaarrekening 2

De tabel toont automatisch alle huidge transacties van die sheet.

### Stap 3: Bereid je CSV bestand voor
Maak een CSV bestand met de volgende kolommen:
- Datum (DD-MM-YYYY of YYYY-MM-DD formaat)
- Naam / Omschrijving
- Af Bij (waarde: "Af" of "Bij")
- Bedrag (getal met . of , scheidingsteen)
- Rekening (correct rekeningnummer voor de sheet)
- Valutadatum (optioneel)
- Mededelingen (optioneel)

**NB: Zet de rekeningnummers correct:**
- Bankrekening: `NL40INGB0002691632`
- Spaarrekening 1: `S 858-17363`
- Spaarrekening 2: `D 130-41072`

### Stap 4: Upload je CSV
Klik op "📁 Selecteer Bestand" en kies je CSV bestand.

De app controleert automatisch:
1. ✓ Datum formaat
2. ✓ Bedrag is numeriek
3. ✓ Af Bij waarde
4. ✓ Rekeningnummer matches sheet
5. ✓ Geen duplicaten

### Stap 5: Controleer de preview
De app toont je:
- Eventuele fouten (rood)
- Verwijderde duplicaten (rood)
- Transacties die worden toegevoegd *in cursief met gele achtergrond*

Controleer alles goed en klik "Akkoord, Toevoegen".

### Stap 6: Bevestiging
Na toevoegen:
- ✓ Automatische backup wordt gemaakt
- ✓ Transacties worden in Excel geplaatst
- ✓ Tabel refresh en toont nieuwe data
- ✓ Statistieken update (aantal records, laatste datum)

## CSV Voorbeeld

```csv
Datum,Naam / Omschrijving,Af Bij,Bedrag,Rekening,Valutadatum,Mededelingen
01-03-2026,Huur Kantoor,Af,1500.00,NL40INGB0002691632,01-03-2026,Huurcontract 2026
02-03-2026,Gift Donatie,Bij,100.00,NL40INGB0002691632,02-03-2026,Anoniem
```

## Veelgestelde vragen

**V: Wat gebeurt er als er fouten in mijn CSV zijn?**
A: De app toont een lijst met alle fouten per rij. Je moet het CSV bestand aanpassen en opnieuw proberen.

**V: Kan ik duplicaten handmatig toevoegen?**
A: Nee, de app weigert duplicaten. Dit ter bescherming van data integriteit.

**V: Waar word mijn backup opgeslagen?**
A: In de folder ingesteld in config.json onder `backup_directory`. Default: `...SharePoint.../backup`

**V: Kan ik de app sluiten zonder gegevens kwijt te raken?**
A: Ja, je kunt op "Afsluiten" klikken. Alleen reeds toegevoegde transacties zijn opgeslagen.

**V: Hoe weet ik hoeveel transacties ik heb toegevoegd?**
A: Na succesvol toevoegen toont de app een groene mededeling met het aantal toegevoegde records. Ook update de statistieken kaarten.

## Data Profiling Regels

Alle transacties worden geverifieerd tegen deze regels:

| Veld | Regel | Voorbeeld |
|------|-------|-----------|
| Datum | Moet geldig zijn | 01-03-2026 |
| Bedrag | Numeriek (. of ,) | 1500.00 of 1500,50 |
| Af Bij | Exact "Af" of "Bij" | Af |
| Rekening | Sheet-specifiek rekeningnummer | NL40INGB0002691632 |

## Problemen?

Controleer het log bestand:
- Windows: `...Betaalde sharepoint.../log/transactietoevoegen_webapp_log.txt`

Of herstart de app en probeer opnieuw.
