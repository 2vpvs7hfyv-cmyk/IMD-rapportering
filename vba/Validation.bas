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
                    cell.Style = "DŒlig"
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
        MsgBox "Verifiering: Hittade " & badCount & " rader med 0 i fšrbrukning i kolumn G. Dessa har markerats.", vbExclamation, "Verifiering - Fšrbrukning"
        VerifyConsumptionZero = False
    End If
End Function

Public Function AllValidationsPass(ByVal targetSheet As Worksheet) As Boolean
    If targetSheet Is Nothing Then Exit Function

    AllValidationsPass = False

    If Not VerifyConsumptionZero(targetSheet) Then Exit Function

    AllValidationsPass = True
End Function
