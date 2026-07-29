# Användning

1. Öppna arbetsboken i Excel på Mac.
2. Gå till VBA-editorn och importera filen `vba/IMDReporting.bas`.
3. Kör makrot `ImportSelectedWorkbookToNewSheet` för att välja en Excel-fil från Finder.
4. Innehållet från den valda filen kopieras till ett nytt blad med samma namn som filen.

Exempel:

```vb
Sub TestImport()
    ImportSelectedWorkbookToNewSheet
End Sub
```
