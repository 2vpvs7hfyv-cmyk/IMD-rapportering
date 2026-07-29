Option Explicit

Public Sub InitializeIMDReport()
    Dim ws As Worksheet
    Dim headers As Variant
    Dim reportName As String

    reportName = "IMD Report"

    Set ws = GetOrCreateWorksheet(reportName)

    headers = Array("Datum", "Avdelning", "Konto", "Värde", "Kommentar")

    With ws
        .Cells.Clear
        .Range("A1:E1").Value = headers
        .Rows(1).Font.Bold = True
        .Range("A1:E1").Borders.Weight = xlThin
        .Columns("A:E").AutoFit
        .Range("A2").Select
    End With
End Sub

Public Sub AppendIMDRow(ByVal department As String, ByVal account As String, ByVal value As Double, Optional ByVal comment As String = "")
    Dim ws As Worksheet
    Dim nextRow As Long

    Set ws = GetOrCreateWorksheet("IMD Report")

    nextRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row + 1

    ws.Cells(nextRow, 1).Value = Date
    ws.Cells(nextRow, 2).Value = department
    ws.Cells(nextRow, 3).Value = account
    ws.Cells(nextRow, 4).Value = value
    ws.Cells(nextRow, 5).Value = comment

    ws.Cells(nextRow, 4).NumberFormat = "#,##0.00"
    ws.Cells(nextRow, 1).NumberFormat = "yyyy-mm-dd"
End Sub

Public Sub ImportSelectedWorkbookToNewSheet()
    Dim selectedFile As Variant
    Dim sourceWorkbook As Workbook
    Dim sourceSheet As Worksheet
    Dim newSheet As Worksheet
    Dim sheetName As String
    Dim filePath As String

    ' Use GetOpenFilename for compatibility on macOS
    selectedFile = Application.GetOpenFilename("Excel Files (*.xlsx;*.xlsm;*.xls), *.xlsx;*.xlsm;*.xls", , "Välj Excel-fil att importera")

    ' If user cancelled, GetOpenFilename returns False (vbBoolean)
    If VarType(selectedFile) = vbBoolean Then
        Exit Sub
    End If

    filePath = CStr(selectedFile)
    sheetName = GetSafeSheetName(GetFileNameWithoutExtension(filePath))

    On Error GoTo ImportError
    Set sourceWorkbook = Workbooks.Open(filePath, UpdateLinks:=False, ReadOnly:=True)

    If sourceWorkbook.Worksheets.Count < 1 Then
        MsgBox "Källarbetsboken innehåller inga blad.", vbExclamation
        sourceWorkbook.Close SaveChanges:=False
        Exit Sub
    End If

    Set sourceSheet = sourceWorkbook.Worksheets(1)

    Set newSheet = GetOrCreateWorksheet(sheetName)
    newSheet.Cells.Clear
    sourceSheet.UsedRange.Copy Destination:=newSheet.Range("A1")

    sourceWorkbook.Close SaveChanges:=False

    MsgBox "Import klar. Innehållet från " & filePath & vbCrLf & "lades till i bladet " & sheetName, vbInformation
    Exit Sub

ImportError:
    If Not sourceWorkbook Is Nothing Then
        On Error Resume Next
        sourceWorkbook.Close SaveChanges:=False
    End If

    MsgBox "Det gick inte att importera filen." & vbCrLf & Err.Description, vbExclamation
End Sub

Private Function GetOrCreateWorksheet(ByVal sheetName As String) As Worksheet
    On Error Resume Next
    Set GetOrCreateWorksheet = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0

    If GetOrCreateWorksheet Is Nothing Then
        Set GetOrCreateWorksheet = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        GetOrCreateWorksheet.Name = sheetName
    End If
End Function

Private Function GetSafeSheetName(ByVal proposedName As String) As String
    Dim safeName As String
    Dim index As Long

    safeName = Trim(proposedName)
    safeName = Replace(safeName, "/", "_")
    safeName = Replace(safeName, "\", "_")
    safeName = Replace(safeName, ":", "_")
    safeName = Replace(safeName, "?", "_")
    safeName = Replace(safeName, "*", "_")
    safeName = Replace(safeName, "[", "_")
    safeName = Replace(safeName, "]", "_")
    safeName = Replace(safeName, " ", "_")

    If Len(safeName) > 31 Then
        safeName = Left$(safeName, 31)
    End If

    If safeName = "" Then
        safeName = "ImportedSheet"
    End If

    If SheetExists(safeName) Then
        index = 1
        Do While SheetExists(safeName & "_" & CStr(index))
            index = index + 1
        Loop
        safeName = safeName & "_" & CStr(index)
    End If

    GetSafeSheetName = safeName
End Function

Private Function GetFileNameWithoutExtension(ByVal fullPath As String) As String
    Dim separatorPosition As Long

    separatorPosition = InStrRev(fullPath, Application.PathSeparator)
    If separatorPosition > 0 Then
        fullPath = Mid$(fullPath, separatorPosition + 1)
    End If

    separatorPosition = InStrRev(fullPath, ".")
    If separatorPosition > 0 Then
        GetFileNameWithoutExtension = Left$(fullPath, separatorPosition - 1)
    Else
        GetFileNameWithoutExtension = fullPath
    End If
End Function

Private Function SheetExists(ByVal sheetName As String) As Boolean
    On Error Resume Next
    SheetExists = Not ThisWorkbook.Worksheets(sheetName) Is Nothing
    On Error GoTo 0
End Function
