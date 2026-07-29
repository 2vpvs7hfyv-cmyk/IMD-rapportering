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
4. Du kan också lägga till en knapp i Excel och koppla den till `ImportSelectedWorkbookToNewSheet`, `ImportLastWorkbookToNewSheet` eller `UpdateVbaModulesFromFolderButton`.

## Git
Detta bibliotek är initierat som ett lokalt Git-repo. För att koppla det till GitHub kan du köra:

```bash
git remote add origin https://github.com/<ditt-användarnamn>/IMD-rapportering.git
```
