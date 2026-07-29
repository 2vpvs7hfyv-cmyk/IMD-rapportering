Option Explicit

Private Sub Workbook_Open()
    On Error Resume Next
    ImportLastWorkbookToNewSheet
    If Err.Number <> 0 Then
        Err.Clear
        MsgBox "Automatisk import misslyckades vid öppning. Välj fil manuellt med ImportSelectedWorkbookToNewSheet.", vbExclamation
    End If
    On Error GoTo 0
End Sub
