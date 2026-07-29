Option Explicit

Public Sub ImportBasModulesFromWorkbookFolder()
    Dim wbFolder As String
    Dim basFile As String
    Dim importCount As Long

    wbFolder = GetWorkbookFolderPath()
    If wbFolder = "" Then Exit Sub

    basFile = Dir(wbFolder & "*.bas")
    Do While basFile <> ""
        ImportBasModule wbFolder & basFile
        importCount = importCount + 1
        basFile = Dir()
    Loop

    If importCount > 0 Then
        MsgBox CStr(importCount) & " .bas-modul(er) importerades från mappen: " & wbFolder, vbInformation
    End If
End Sub

Public Sub UpdateVbaModulesFromFolderButton()
    ImportBasModulesFromWorkbookFolder
End Sub

Private Sub ImportBasModule(ByVal filePath As String)
    Dim vbComp As Object

    On Error Resume Next
    Set vbComp = ThisWorkbook.VBProject.VBComponents.Import(filePath)
    If Err.Number <> 0 Then
        Err.Clear
        MsgBox "Kunde inte importera modul: " & filePath, vbExclamation
    End If
    On Error GoTo 0
End Sub

Private Function GetWorkbookFolderPath() As String
    If ThisWorkbook.Path = "" Then
        GetWorkbookFolderPath = CurDir()
    Else
        GetWorkbookFolderPath = ThisWorkbook.Path & Application.PathSeparator
    End If
End Function
