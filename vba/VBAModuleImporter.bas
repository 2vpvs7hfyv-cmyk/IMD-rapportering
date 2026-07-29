Option Explicit

Public Sub ImportBasModulesFromWorkbookFolder()
    Dim wbFolder As String
    Dim basFile As String
    Dim importCount As Long

    wbFolder = GetWorkbookFolderPath()
    If wbFolder = "" Then Exit Sub

    RemoveAllStandardModules

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

Private Sub RemoveAllStandardModules()
    Dim comp As Object
    Dim moduleNames As Collection
    Dim i As Long

    Set moduleNames = New Collection

    For Each comp In ThisWorkbook.VBProject.VBComponents
        If comp.Type = 1 Then
            moduleNames.Add comp.Name
        End If
    Next comp

    For i = 1 To moduleNames.Count
        ThisWorkbook.VBProject.VBComponents.Remove ThisWorkbook.VBProject.VBComponents(moduleNames(i))
    Next i
End Sub

Public Sub UpdateVbaModulesFromFolderButton()
    ImportBasModulesFromWorkbookFolder
End Sub

Private Sub ImportBasModule(ByVal filePath As String)
    Dim moduleName As String
    Dim existingComp As Object
    Dim vbComp As Object

    moduleName = GetModuleNameFromBasPath(filePath)

    On Error Resume Next
    Set existingComp = ThisWorkbook.VBProject.VBComponents(moduleName)
    If Not existingComp Is Nothing Then
        ThisWorkbook.VBProject.VBComponents.Remove existingComp
        Set existingComp = Nothing
    End If
    Err.Clear

    Set vbComp = ThisWorkbook.VBProject.VBComponents.Import(filePath)
    If Err.Number <> 0 Then
        Err.Clear
        MsgBox "Kunde inte importera modul: " & filePath, vbExclamation
    End If
    On Error GoTo 0
End Sub

Private Function GetModuleNameFromBasPath(ByVal filePath As String) As String
    Dim fileName As String
    Dim lastSlash As Long
    Dim lastBackslash As Long
    Dim lastColon As Long
    Dim lastSep As Long

    lastSlash = InStrRev(filePath, "/")
    lastBackslash = InStrRev(filePath, "\")
    lastColon = InStrRev(filePath, ":")
    lastSep = lastSlash
    If lastBackslash > lastSep Then lastSep = lastBackslash
    If lastColon > lastSep Then lastSep = lastColon

    If lastSep > 0 Then
        fileName = Mid$(filePath, lastSep + 1)
    Else
        fileName = filePath
    End If

    If InStrRev(fileName, ".") > 0 Then
        GetModuleNameFromBasPath = Left$(fileName, InStrRev(fileName, ".") - 1)
    Else
        GetModuleNameFromBasPath = fileName
    End If
End Function

Private Function GetWorkbookFolderPath() As String
    If ThisWorkbook.Path = "" Then
        GetWorkbookFolderPath = CurDir()
    Else
        GetWorkbookFolderPath = ThisWorkbook.Path & Application.PathSeparator
    End If
End Function
