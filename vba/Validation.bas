Option Explicit

Public Function VerifyConsumptionZero(ByVal targetSheet As Worksheet) As Boolean
    Dim lastRow As Long
    Dim checkRange As Range
    Dim cell As Range
    Dim badCount As Long

    VerifyConsumptionZero = True
    If targetSheet Is Nothing Then Exit Function

    lastRow = targetSheet.Cells(targetSheet.Rows.Count, "G").End(xlUp).Row
    If lastRow < 2 Then Exit Function

    Set checkRange = targetSheet.Range("G2:G" & lastRow)
    badCount = 0

    For Each cell In checkRange
        If Not IsError(cell.Value) Then
            If Trim$(CStr(cell.Value)) <> "" Then
                If Val(cell.Value) = 0 Then
                    On Error Resume Next
                    cell.Style = "Dålig"
                    If Err.Number <> 0 Then
                        Err.Clear
                        cell.Interior.Color = RGB(255, 199, 206)
                    End If
                    On Error GoTo 0
                    badCount = badCount + 1
                End If
            End If
        End If
    Next cell

    If badCount > 0 Then
        MsgBox "Verifiering: Hittade " & badCount & " rader med 0 i förbrukning i kolumn G. Dessa har markerats.", vbExclamation, "Verifiering - Förbrukning"
        VerifyConsumptionZero = False
    End If
End Function

Public Function AllValidationsPass(ByVal targetSheet As Worksheet) As Boolean
    If targetSheet Is Nothing Then Exit Function

    AllValidationsPass = False

    If Not VerifyConsumptionZero(targetSheet) Then Exit Function

    AllValidationsPass = True
End Function

Private Function ParseConsumptionValue(ByVal cellValue As Variant) As Double
    If IsError(cellValue) Then
        ParseConsumptionValue = 0
        Exit Function
    End If
    If IsNumeric(cellValue) Then
        ParseConsumptionValue = CDbl(cellValue)
        Exit Function
    End If
    Dim textVal As String
    textVal = Trim$(CStr(cellValue))
    textVal = Replace(textVal, " ", "")
    If InStr(textVal, ",") > 0 Then
        textVal = Replace(textVal, ".", "")
        textVal = Replace(textVal, ",", ".")
    End If
    If IsNumeric(textVal) Then
        ParseConsumptionValue = CDbl(textVal)
    End If
End Function

Private Function GetYearOverYearThreshold() As Double
    On Error Resume Next
    Dim val As Variant
    val = ThisWorkbook.Names("Förbrukningströskel").RefersToRange.Value
    If Err.Number = 0 Then
        On Error GoTo 0
        If IsNumeric(val) And Trim$(CStr(val)) <> "" Then
            GetYearOverYearThreshold = CDbl(val)
        Else
            GetYearOverYearThreshold = 0
        End If
        Exit Function
    End If
    Err.Clear
    On Error GoTo 0
    GetYearOverYearThreshold = 0
End Function

Private Function GetPreviousYearSuffix(ByVal currentSuffix As String) As String
    If Len(currentSuffix) <> 4 Then Exit Function
    If Not IsNumeric(currentSuffix) Then Exit Function
    Dim yy As Integer
    Dim mm As String
    yy = CInt(Left$(currentSuffix, 2))
    mm = Right$(currentSuffix, 2)
    GetPreviousYearSuffix = Format(yy - 1, "00") & mm
End Function

Private Function FindSheetBySuffix(ByVal suffix As String) As Worksheet
    Dim ws As Worksheet
    Set FindSheetBySuffix = Nothing
    For Each ws In ThisWorkbook.Worksheets
        If Right$(ws.Name, Len(suffix)) = suffix Then
            Set FindSheetBySuffix = ws
            Exit Function
        End If
    Next ws
End Function

Public Function VerifyYearOverYearConsumption(ByVal targetSheet As Worksheet, ByRef statusMessage As String) As Boolean
    VerifyYearOverYearConsumption = True
    statusMessage = ""

    If targetSheet Is Nothing Then Exit Function

    Dim currentSuffix As String
    Dim prevSuffix As String
    Dim prevSheet As Worksheet
    Dim threshold As Double
    Dim lastRow As Long
    Dim prevLastRow As Long
    Dim row As Long
    Dim searchRow As Long
    Dim diffCount As Long
    Dim meterID As String
    Dim currentConsumption As Double
    Dim prevRow As Long
    Dim prevConsumption As Double

    If Len(targetSheet.Name) >= 4 Then
        currentSuffix = Right$(targetSheet.Name, 4)
    End If

    If Not IsNumeric(currentSuffix) Then
        statusMessage = "Årsjämförelse: Bladnamnet slutar inte med YYMM - jämförelse hoppades över."
        Exit Function
    End If

    prevSuffix = GetPreviousYearSuffix(currentSuffix)

    Set prevSheet = FindSheetBySuffix(prevSuffix)
    If prevSheet Is Nothing Then
        statusMessage = "Årsjämförelse: Inget blad med suffix " & prevSuffix & " hittades - jämförelse hoppades över."
        Exit Function
    End If

    threshold = GetYearOverYearThreshold()
    diffCount = 0
    lastRow = targetSheet.Cells(targetSheet.Rows.Count, "A").End(xlUp).Row
    prevLastRow = prevSheet.Cells(prevSheet.Rows.Count, "A").End(xlUp).Row

    For row = 2 To lastRow
        meterID = Trim$(CStr(targetSheet.Cells(row, "A").Value))
        If meterID <> "" Then
            currentConsumption = ParseConsumptionValue(targetSheet.Cells(row, "G").Value)
            prevRow = 0
            For searchRow = 2 To prevLastRow
                If Trim$(CStr(prevSheet.Cells(searchRow, "A").Value)) = meterID Then
                    prevRow = searchRow
                    Exit For
                End If
            Next searchRow
            If prevRow > 0 Then
                prevConsumption = ParseConsumptionValue(prevSheet.Cells(prevRow, "G").Value)
                If Abs(currentConsumption - prevConsumption) > threshold Then
                    On Error Resume Next
                    targetSheet.Cells(row, "G").Style = "Dålig"
                    If Err.Number <> 0 Then
                        Err.Clear
                        targetSheet.Cells(row, "G").Interior.Color = RGB(255, 199, 206)
                    End If
                    On Error GoTo 0
                    diffCount = diffCount + 1
                End If
            End If
        End If
    Next row

    If diffCount > 0 Then
        statusMessage = "Årsjämförelse (" & prevSheet.Name & "): " & diffCount & " mätare avviker mer än " & threshold & " kWh från föregående år. Avvikelserna är markerade med rött i bladet."
        VerifyYearOverYearConsumption = False
    Else
        statusMessage = "Årsjämförelse (" & prevSheet.Name & "): Alla mätare inom acceptabel variation (tröskel: " & threshold & " kWh)."
    End If
End Function
