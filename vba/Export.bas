Option Explicit

Public Sub BuildExportSheet()
    Dim sheetName As String
    Dim exportSheet As Worksheet
    Dim exportedCount As Long

    sheetName = GetLastImportedSheetName()
    If Trim$(sheetName) = "" Then
        MsgBox "Ingen importerad fil hittades att exportera. Kšr importen fšrst.", vbExclamation
        Exit Sub
    End If

    If Not GetLastImportValidationStatus() Then
        MsgBox "Senaste importen har inte validerats eller valideringen misslyckades. Kšr importen igen och se till att den godkŠnns innan export.", vbExclamation
        Exit Sub
    End If

    If Not SheetExists(sheetName) Then
        MsgBox "Det importerade bladet '" & sheetName & "' finns inte. Kšr importen igen.", vbExclamation
        Exit Sub
    End If

    Set exportSheet = ThisWorkbook.Worksheets(sheetName)
    exportedCount = TransformDataToExportFormat(exportSheet)
    If exportedCount > 0 Then
        MsgBox "Export klar. InnehŒllet frŒn bladet '" & sheetName & "' exporterades till bladet Export." & vbCrLf & "Antal elmŠtare: " & exportedCount, vbInformation
    Else
        MsgBox "Exporten kšrdes, men inga rader kunde exporteras.", vbExclamation
    End If
End Sub


Public Function TransformDataToExportFormat(ByVal sourceSheet As Worksheet) As Long
    Dim exportSheet As Worksheet
    Dim sourceLastRow As Long
    Dim sourceRow As Long
    Dim destRow As Long
    Dim sourceA As String
    Dim sourceD As String
    Dim sourceE As Variant
    Dim sourceF As Variant
    Dim sourceG As Variant
    Dim exportTimestamp As String
    Dim exportValueD As String
    Dim exportValueF As String
    Dim exportValueG As String
    Dim exportValueH As String
    Dim exportValueI As String
    Dim exportedCount As Long

    TransformDataToExportFormat = 0
    If sourceSheet Is Nothing Then Exit Function

    On Error Resume Next
    Application.DisplayAlerts = False
    ThisWorkbook.Worksheets("Export").Delete
    Application.DisplayAlerts = True
    On Error GoTo 0

    Set exportSheet = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
    exportSheet.Name = "Export"

    exportTimestamp = Format(Now, "yyyy-mm-dd hh:nn:ss")
    With exportSheet.Range("B1")
        .NumberFormat = "@"
        .Value = "'" & exportTimestamp
    End With

    sourceLastRow = sourceSheet.Cells(sourceSheet.Rows.Count, "A").End(xlUp).Row
    destRow = 2

    For sourceRow = 2 To sourceLastRow
        sourceA = Trim$(CStr(sourceSheet.Cells(sourceRow, "A").Value))
        sourceD = Trim$(CStr(sourceSheet.Cells(sourceRow, "D").Value))
        sourceE = sourceSheet.Cells(sourceRow, "E").Value
        sourceF = sourceSheet.Cells(sourceRow, "F").Value
        sourceG = sourceSheet.Cells(sourceRow, "G").Value

        exportValueD = "213424" & ExtractLast4Digits(sourceA)
        exportValueF = FormatDecimalText(ParseDecimalValue(sourceD) - ParseDecimalValue(sourceG))
        exportValueG = FormatDateToText(sourceF)
        exportValueH = FormatDecimalText(ParseDecimalValue(sourceD))
        exportValueI = FormatDecimalText(ParseDecimalValue(sourceG))

        exportSheet.Cells(destRow, "A").Value = ""
        exportSheet.Cells(destRow, "B").Value = ""
        exportSheet.Cells(destRow, "C").Value = "EL"
        exportSheet.Cells(destRow, "D").Value = exportValueD
        exportSheet.Cells(destRow, "E").Value = FormatDateToText(sourceE)
        exportSheet.Cells(destRow, "F").Value = exportValueF
        exportSheet.Cells(destRow, "G").Value = exportValueG
        exportSheet.Cells(destRow, "H").Value = exportValueH
        exportSheet.Cells(destRow, "I").Value = exportValueI
        exportSheet.Cells(destRow, "J").Value = "2"
        exportSheet.Cells(destRow, "K").Value = "2,5"
        exportSheet.Cells(destRow, "L").Value = ""
        exportSheet.Cells(destRow, "M").Value = ""
        exportSheet.Cells(destRow, "N").Value = ""
        exportSheet.Cells(destRow, "O").Value = ""
        exportSheet.Cells(destRow, "P").Value = "kWh"
        exportSheet.Cells(destRow, "Q").Value = "Paviljongen IMD"

        exportSheet.Range(exportSheet.Cells(destRow, 1), exportSheet.Cells(destRow, 17)).NumberFormat = "@"
        destRow = destRow + 1
        exportedCount = exportedCount + 1
    Next sourceRow

    TransformDataToExportFormat = exportedCount
End Function

Private Function ExtractLast4Digits(ByVal inputText As String) As String
    Dim digits As String
    Dim i As Long
    Dim ch As String

    digits = ""
    For i = 1 To Len(inputText)
        ch = Mid$(inputText, i, 1)
        If ch >= "0" And ch <= "9" Then
            digits = digits & ch
        End If
    Next i

    If Len(digits) >= 4 Then
        ExtractLast4Digits = Right$(digits, 4)
    Else
        ExtractLast4Digits = digits
    End If
End Function

Private Function ParseDecimalValue(ByVal sourceValue As Variant) As Double
    Dim textValue As String

    If IsError(sourceValue) Then
        ParseDecimalValue = 0
        Exit Function
    End If

    If IsNumeric(sourceValue) Then
        ParseDecimalValue = CDbl(sourceValue)
        Exit Function
    End If

    textValue = Trim$(CStr(sourceValue))
    If textValue = "" Then
        ParseDecimalValue = 0
        Exit Function
    End If

    textValue = Replace(textValue, " ", "")
    If InStr(textValue, ",") > 0 Then
        textValue = Replace(textValue, ".", "")
        textValue = Replace(textValue, ",", ".")
    End If

    If IsNumeric(textValue) Then
        ParseDecimalValue = CDbl(textValue)
    Else
        ParseDecimalValue = 0
    End If
End Function

Private Function FormatDecimalText(ByVal value As Double) As String
    Dim formatted As String

    formatted = Format(value, "0.############")
    formatted = Replace(formatted, ".", ",")
    If Right$(formatted, 1) = "," Then formatted = Left$(formatted, Len(formatted) - 1)
    If formatted = "-0" Then formatted = "0"
    FormatDecimalText = formatted
End Function

Private Function FormatDateToText(ByVal sourceValue As Variant) As String
    Dim result As String

    If IsError(sourceValue) Then
        FormatDateToText = ""
        Exit Function
    End If

    If IsDate(sourceValue) Then
        result = Format(CDate(sourceValue), "yyyymmdd")
        FormatDateToText = result
        Exit Function
    End If

    result = Trim$(CStr(sourceValue))
    If result = "" Then
        FormatDateToText = ""
        Exit Function
    End If

    On Error Resume Next
    result = Format(CDate(result), "yyyymmdd")
    If Err.Number <> 0 Then
        Err.Clear
        result = Trim$(CStr(sourceValue))
    End If
    On Error GoTo 0

    FormatDateToText = result
End Function

Public Sub SaveExportSheetAsCSV()
    Dim exportSheet As Worksheet
    Dim filePath As String
    Dim fileNumber As Integer
    Dim lastRow As Long
    Dim currentRow As Long
    Dim lastCol As Long
    Dim currentCol As Long
    Dim csvLine As String
    Dim cellValue As String
    Dim wbFolder As String

    On Error GoTo ErrorHandler

    Set exportSheet = Nothing
    On Error Resume Next
    Set exportSheet = ThisWorkbook.Worksheets("Export")
    On Error GoTo ErrorHandler

    If exportSheet Is Nothing Then
        MsgBox "Export-bladet finns inte. Kï¿½r fï¿½rst en import och validering.", vbExclamation, "Fel"
        Exit Sub
    End If

    lastRow = exportSheet.Cells(exportSheet.Rows.Count, 3).End(xlUp).Row
    If lastRow < 2 Then
        MsgBox "Export-bladet ï¿½r tomt.", vbInformation, "Ingen data"
        Exit Sub
    End If

    wbFolder = GetWorkbookFolderPath()
    If wbFolder = "" Then
        MsgBox "Kunde inte avgï¿½ra arbetsbokens mapp.", vbExclamation, "Fel"
        Exit Sub
    End If

    filePath = wbFolder & "export_" & Format(Now, "yyyymmdd_hhmmss") & ".txt"

    fileNumber = FreeFile
    Open filePath For Output As fileNumber

    For currentRow = 1 To lastRow
        csvLine = ""
        For currentCol = 1 To 17
            cellValue = Trim$(CStr(exportSheet.Cells(currentRow, currentCol).Value))
            If currentCol > 1 Then
                csvLine = csvLine & ";"
            End If
            csvLine = csvLine & cellValue
        Next currentCol
        Print #fileNumber, csvLine
    Next currentRow

    Close fileNumber

    MsgBox "Export-data sparad till:" & vbCrLf & filePath, vbInformation, "Export klar"
    Exit Sub

ErrorHandler:
    If fileNumber <> 0 Then
        Close fileNumber
    End If
    MsgBox "Fel vid export: " & Err.Description, vbExclamation, "Exportfel"
End Sub

Public Function GetWorkbookFolderPath() As String
    If ThisWorkbook.Path = "" Then
        GetWorkbookFolderPath = CurDir()
    Else
        GetWorkbookFolderPath = ThisWorkbook.Path & Application.PathSeparator
    End If
End Function
