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
    sheetName = GetSafeSheetName(GetFileNameWithoutExtension(filePath))

    On Error GoTo ImportError
    Set sourceWorkbook = Workbooks.Open(filePath, UpdateLinks:=False, ReadOnly:=True)

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
