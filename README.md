# IMD-rapportering
Excel VBA project for IMD reporting

Detta är ett starterprojekt för Excel VBA med Git-versionering.

## Struktur
- `vba/` – VBA-moduler och makron
- `docs/` – dokumentation och anteckningar

## Kom igång
1. Öppna Excel och skapa en arbetsbok.
2. Lägg till VBA-modulen från mappen `vba/`.
3. I VBA-editorn, kör `ImportSelectedWorkbookToNewSheet` för att testa importfunktionen.
4. Lägg gärna till två knappar i Excel:
   - `ImportSelectedWorkbookToNewSheet` för import + validering
   - `ExportLastImportedSheet` för att exportera det senaste validerade bladet
5. Om du vill uppdatera VBA-koden från mappen, lägg till en knapp som kör `UpdateVbaModulesFromFolderButton`.

## Git
Detta bibliotek är initierat som ett lokalt Git-repo. För att koppla det till GitHub kan du köra:

```bash
git remote add origin https://github.com/<ditt-användarnamn>/IMD-rapportering.git
```
