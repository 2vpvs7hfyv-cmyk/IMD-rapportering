Option Explicit

'' Public helper routines were archived to vba/archive/IMDReporting_legacy.bas
'' to keep the active module focused on import functionality.

Public Sub ImportSelectedWorkbookToNewSheet()
    Dim selectedFile As Variant
    Dim sourceWorkbook As Workbook
    Dim sourceSheet As Worksheet
    Dim newSheet As Worksheet
    Dim sheetName As String
    Dim filePath As String

    ' Use GetOpenFilename for compatibility on macOS
    Dim pickFailed As Boolean
    pickFailed = False

    On Error Resume Next
    selectedFile = Application.GetOpenFilename("Excel Files (*.xlsx;*.xlsm;*.xls), *.xlsx;*.xlsm;*.xls", , "Välj Excel-fil att importera")
    If Err.Number <> 0 Then
        pickFailed = True
        Err.Clear
    End If
    On Error GoTo 0

    ' If GetOpenFilename returned False or failed, try AppleScript fallback on Mac
    If VarType(selectedFile) = vbBoolean Or pickFailed Then
        On Error Resume Next
        Dim asPath As String
        asPath = MacScript("POSIX path of (choose file with prompt ""Välj Excel-fil att importera"")")
        If Err.Number <> 0 Then
            Err.Clear
        Else
            If Len(Trim$(asPath)) > 0 Then
                selectedFile = asPath
            End If
        End If
        On Error GoTo 0
    End If

    ' If still no file selected, ask the user to paste the full POSIX path
    If VarType(selectedFile) = vbBoolean Or Trim(CStr(selectedFile)) = "" Then
        Dim inputPath As String
        inputPath = InputBox("Ange full sökväg till Excel-filen (kopiera sökvägen från Finder och klistra in):", "Ange filväg")
        If Trim(inputPath) = "" Then Exit Sub
        selectedFile = inputPath
    End If

    filePath = CStr(selectedFile)
    ImportWorkbookToNewSheetByPath filePath
End Sub

Public Sub ImportLastWorkbookToNewSheet()
    If Not TryImportLastWorkbookToNewSheet() Then
        MsgBox "Ingen tidigare importfil är sparad eller filen kunde inte hittas. Välj en fil att importera.", vbInformation
        ImportSelectedWorkbookToNewSheet
    End If
End Sub

Public Function TryImportLastWorkbookToNewSheet() As Boolean
    Dim lastPath As String
    lastPath = GetLastImportedFilePath()

    If Trim$(lastPath) = "" Then Exit Function
    If Dir(lastPath) = "" Then Exit Function

    ImportWorkbookToNewSheetByPath lastPath
    TryImportLastWorkbookToNewSheet = True
End Function

Public Sub ImportWorkbookToNewSheet(ByVal filePath As String)
    If Trim$(CStr(filePath)) = "" Then
        MsgBox "Ogiltig filväg.", vbExclamation
        Exit Sub
    End If

    ImportWorkbookToNewSheetByPath CStr(filePath)
End Sub

Private Sub ImportWorkbookToNewSheetByPath(ByVal filePath As String)
    Dim sourceWorkbook As Workbook
    Dim sourceSheet As Worksheet
    Dim newSheet As Worksheet
    Dim sheetName As String

    sheetName = GetSafeSheetName(GetFileNameWithoutExtension(filePath))

    On Error GoTo ImportError
    Set sourceWorkbook = Workbooks.Open(filePath, UpdateLinks:=False, ReadOnly:=True)

    SaveLastImportedFilePath filePath

    If sourceWorkbook.Worksheets.Count < 1 Then
        MsgBox "Källarbetsboken innehåller inga blad.", vbExclamation
        sourceWorkbook.Close SaveChanges:=False
        Exit Sub
    End If

    Set sourceSheet = sourceWorkbook.Worksheets(1)

    If SheetExists(sheetName) Then
        Dim overwriteResponse As VbMsgBoxResult
        overwriteResponse = MsgBox("Bladet '" & sheetName & "' finns redan. Vill du skriva över det?", vbYesNo + vbQuestion, "Skriv över befintligt blad?")
        If overwriteResponse = vbNo Then
            sourceWorkbook.Close SaveChanges:=False
            Exit Sub
        End If
        Set newSheet = ThisWorkbook.Worksheets(sheetName)
        newSheet.Cells.Clear
    Else
        Set newSheet = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        newSheet.Name = sheetName
    End If

    sourceSheet.UsedRange.Copy Destination:=newSheet.Range("A1")

    If AllValidationsPass(newSheet) Then
        SaveLastImportedSheetName sheetName
        SaveLastImportValidationStatus True
        sourceWorkbook.Close SaveChanges:=False
        MsgBox "Import klar. Innehållet från " & filePath & vbCrLf & "lades till i bladet " & sheetName & ". Valideringen lyckades. Kör export separat med exportknappen.", vbInformation
    Else
        SaveLastImportedSheetName sheetName
        SaveLastImportValidationStatus False
        sourceWorkbook.Close SaveChanges:=False
        MsgBox "Importen genomfördes, men verifieringarna misslyckades. Export körs inte.", vbExclamation
    End If

    Exit Sub

ImportError:
    If Not sourceWorkbook Is Nothing Then
        On Error Resume Next
        sourceWorkbook.Close SaveChanges:=False
    End If

    MsgBox "Det gick inte att importera filen." & vbCrLf & Err.Description, vbExclamation
End Sub

Private Sub SaveLastImportedFilePath(ByVal filePath As String)
    On Error Resume Next
    ThisWorkbook.Names("LastImportedFilePath").Delete
    On Error GoTo 0

    ThisWorkbook.Names.Add Name:="LastImportedFilePath", RefersTo:="=""" & Replace(filePath, """", """""") & """"
End Sub

Private Function GetLastImportedFilePath() As String
    On Error Resume Next
    GetLastImportedFilePath = ThisWorkbook.Names("LastImportedFilePath").RefersTo
    If Err.Number <> 0 Then
        Err.Clear
        GetLastImportedFilePath = ""
    Else
        If Left$(GetLastImportedFilePath, 1) = "=" Then
            GetLastImportedFilePath = Mid$(GetLastImportedFilePath, 2)
        End If
        If Left$(GetLastImportedFilePath, 1) = """" Then
            GetLastImportedFilePath = Mid$(GetLastImportedFilePath, 2, Len(GetLastImportedFilePath) - 2)
        End If
    End If
    On Error GoTo 0
End Function

Private Sub SaveLastImportedSheetName(ByVal sheetName As String)
    On Error Resume Next
    ThisWorkbook.Names("LastImportedSheetName").Delete
    On Error GoTo 0

    ThisWorkbook.Names.Add Name:="LastImportedSheetName", RefersTo:="=""" & Replace(sheetName, """", """""") & """"
End Sub

Private Function GetLastImportedSheetName() As String
    On Error Resume Next
    GetLastImportedSheetName = ThisWorkbook.Names("LastImportedSheetName").RefersTo
    If Err.Number <> 0 Then
        Err.Clear
        GetLastImportedSheetName = ""
    Else
        If Left$(GetLastImportedSheetName, 1) = "=" Then
            GetLastImportedSheetName = Mid$(GetLastImportedSheetName, 2)
        End If
        If Left$(GetLastImportedSheetName, 1) = """" Then
            GetLastImportedSheetName = Mid$(GetLastImportedSheetName, 2, Len(GetLastImportedSheetName) - 2)
        End If
    End If
    On Error GoTo 0
End Function

Private Sub SaveLastImportValidationStatus(ByVal status As Boolean)
    On Error Resume Next
    ThisWorkbook.Names("LastImportValidated").Delete
    On Error GoTo 0

    ThisWorkbook.Names.Add Name:="LastImportValidated", RefersTo:="=" & IIf(status, "TRUE", "FALSE")
End Sub

Private Function GetLastImportValidationStatus() As Boolean
    Dim rawValue As String
    On Error Resume Next
    rawValue = ThisWorkbook.Names("LastImportValidated").RefersTo
    If Err.Number <> 0 Then
        Err.Clear
        GetLastImportValidationStatus = False
    Else
        If Left$(rawValue, 1) = "=" Then rawValue = Mid$(rawValue, 2)
        rawValue = UCase$(Trim$(rawValue))
        GetLastImportValidationStatus = (rawValue = "TRUE")
    End If
    On Error GoTo 0
End Function

Public Sub ExportLastImportedSheet()
    Dim sheetName As String
    Dim exportSheet As Worksheet
    Dim exportedCount As Long

    sheetName = GetLastImportedSheetName()
    If Trim$(sheetName) = "" Then
        MsgBox "Ingen importerad fil hittades att exportera. Kör importen först.", vbExclamation
        Exit Sub
    End If

    If Not GetLastImportValidationStatus() Then
        MsgBox "Senaste importen har inte validerats eller valideringen misslyckades. Kör importen igen och se till att den godkänns innan export.", vbExclamation
        Exit Sub
    End If

    If Not SheetExists(sheetName) Then
        MsgBox "Det importerade bladet '" & sheetName & "' finns inte. Kör importen igen.", vbExclamation
        Exit Sub
    End If

    Set exportSheet = ThisWorkbook.Worksheets(sheetName)
    exportedCount = CreateExportSheet(exportSheet)
    If exportedCount > 0 Then
        MsgBox "Export klar. Innehållet från bladet '" & sheetName & "' exporterades till bladet Export." & vbCrLf & "Antal elmätare: " & exportedCount, vbInformation
    Else
        MsgBox "Exporten kördes, men inga rader kunde exporteras.", vbExclamation
    End If
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
    Dim ws As Worksheet

    SheetExists = False
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name = sheetName Then
            SheetExists = True
            Exit Function
        End If
    Next ws
End Function
