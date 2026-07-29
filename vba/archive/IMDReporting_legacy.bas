Option Explicit

' Archived legacy routines from IMDReporting.bas
' These were moved out of the active module to keep the codebase tidy.

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
