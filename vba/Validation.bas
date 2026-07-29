Option Explicit

Public Sub VerifyConsumptionZero(ByVal targetSheet As Worksheet)
    Dim lastRow As Long
    Dim checkRange As Range
    Dim cell As Range
    Dim badCount As Long

    If targetSheet Is Nothing Then Exit Sub

    lastRow = targetSheet.Cells(targetSheet.Rows.Count, "G").End(xlUp).Row
    If lastRow < 2 Then Exit Sub

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
    End If
End Sub
