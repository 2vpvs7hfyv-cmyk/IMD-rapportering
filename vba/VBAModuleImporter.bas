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
    Dim tempFilePath As String
    Dim fileSystem As Object

    moduleName = GetModuleNameFromBasPath(filePath)

    On Error Resume Next
    Set existingComp = ThisWorkbook.VBProject.VBComponents(moduleName)
    If Not existingComp Is Nothing Then
        ThisWorkbook.VBProject.VBComponents.Remove existingComp
        Set existingComp = Nothing
    End If
    Err.Clear

    ' Convert UTF-8 file to MacRoman temporarily for proper import
    Set fileSystem = CreateObject("Scripting.FileSystemObject")
    tempFilePath = fileSystem.GetSpecialFolder(2) & "\import_temp_" & CDbl(Now()) * 100000 & ".bas"
    
    If ConvertFileEncoding(filePath, tempFilePath, "UTF-8", "macRoman") Then
        Set vbComp = ThisWorkbook.VBProject.VBComponents.Import(tempFilePath)
        On Error Resume Next
        fileSystem.DeleteFile tempFilePath
        On Error GoTo 0
    Else
        Set vbComp = ThisWorkbook.VBProject.VBComponents.Import(filePath)
    End If
    
    If Err.Number <> 0 Then
        Err.Clear
        MsgBox "Kunde inte importera modul: " & filePath, vbExclamation
    End If
    On Error GoTo 0
End Sub

Private Function ConvertFileEncoding(sourceFile As String, destFile As String, fromEncoding As String, toEncoding As String) As Boolean
    On Error GoTo ErrorHandler
    Dim stream As Object
    Dim content As String
    
    Set stream = CreateObject("ADODB.Stream")
    
    ' Read as UTF-8
    With stream
        .Charset = "utf-8"
        .Open
        .LoadFromFile sourceFile
        content = .ReadText()
        .Close
    End With
    
    ' Write as MacRoman
    With stream
        .Charset = "macRoman"
        .Open
        .WriteText content
        .SaveToFile destFile, 2
        .Close
    End With
    
    ConvertFileEncoding = True
    Exit Function
    
ErrorHandler:
    ConvertFileEncoding = False
End Function

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

Public Function GetWorkbookFolderPath() As String
    If ThisWorkbook.Path = "" Then
        GetWorkbookFolderPath = CurDir()
    Else
        GetWorkbookFolderPath = ThisWorkbook.Path & Application.PathSeparator
    End If
End Function
