Option Explicit

Public Sub CreateExportSheet(ByVal sourceSheet As Worksheet)
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

    If sourceSheet Is Nothing Then Exit Sub

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

    MsgBox "Export klar. " & exportedCount & " Elmätare exporterade.", vbInformation, "Export klar"
End Sub

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
