Option Explicit

Const pjDoNotSave = 0

Dim mppPath
Dim projectApp
Dim projectFile

Info "Running script: " & WScript.ScriptFullName

If WScript.Arguments.Count < 1 Then
    WScript.Echo "Usage: cscript //nologo open-mpp-file.vbs ""C:\path\to\project.mpp"""
    WScript.Quit 1
End If

mppPath = WScript.Arguments(0)

If Not FileExists(mppPath) Then
    WScript.Echo "MPP file was not found: " & mppPath
    WScript.Quit 1
End If

On Error Resume Next
Set projectApp = GetObject(, "MSProject.Application")
If Err.Number <> 0 Then
    Err.Clear
    Set projectApp = CreateObject("MSProject.Application")
End If

If Err.Number <> 0 Or projectApp Is Nothing Then
    WScript.Echo "Could not start Microsoft Project. Error " & Err.Number & ": " & Err.Description
    WScript.Quit 1
End If
On Error GoTo 0

projectApp.Visible = True

Info "Opening MPP file: " & mppPath
Info "MPP file size: " & GetFileSize(mppPath) & " bytes"

On Error Resume Next
projectApp.FileOpen mppPath
If Err.Number <> 0 Then
    WScript.Echo "Could not open MPP file. Error " & Err.Number & ": " & Err.Description
    WScript.Quit 1
End If
On Error GoTo 0

Info "MPP file opened:  " & mppPath

Set projectFile = projectApp.ActiveProject
EnumerateProjectProperties projectFile
WScript.Echo "------------------"
CloseActiveProjectWithoutSaving projectApp

Sub Info(ByVal message)
    WScript.Echo "info: " & message
End Sub

Function FileExists(ByVal filePath)
    Dim fileSystem
    Set fileSystem = CreateObject("Scripting.FileSystemObject")
    FileExists = fileSystem.FileExists(filePath)
End Function

Function GetFileSize(ByVal filePath)
    Dim fileSystem
    Set fileSystem = CreateObject("Scripting.FileSystemObject")
    GetFileSize = fileSystem.GetFile(filePath).Size
End Function

Sub CloseActiveProjectWithoutSaving(ByVal projectApp)
    On Error Resume Next
    projectApp.FileClose pjDoNotSave
    If Err.Number <> 0 Then
        WScript.Echo "Could not close MPP file without saving. Error " & Err.Number & ": " & Err.Description
        Err.Clear
    Else
        Info "Closed MPP file without saving changes."
    End If
    On Error GoTo 0
End Sub

Sub EnumerateProjectProperties(ByVal projectFile)
    WScript.Echo ""
    WScript.Echo "Project properties"
    WScript.Echo "------------------"

    EchoNamedProjectProperty projectFile, "Name"
    EchoNamedProjectProperty projectFile, "FullName"
    EchoNamedProjectProperty projectFile, "Path"
    EchoNamedProjectProperty projectFile, "Title"
    EchoNamedProjectProperty projectFile, "Subject"
    EchoNamedProjectProperty projectFile, "Author"
    EchoNamedProjectProperty projectFile, "Manager"
    EchoNamedProjectProperty projectFile, "Company"
    EchoNamedProjectProperty projectFile, "Comments"
    EchoNamedProjectProperty projectFile, "CreationDate"
    EchoNamedProjectProperty projectFile, "LastSaveDate"
    EchoNamedProjectProperty projectFile, "Start"
    EchoNamedProjectProperty projectFile, "Finish"
    EchoNamedProjectProperty projectFile, "StatusDate"
    EchoNamedProjectProperty projectFile, "CurrentDate"
    EchoNamedProjectProperty projectFile, "Calendar"
    EchoNamedProjectProperty projectFile, "CurrencySymbol"
    EchoNamedProjectProperty projectFile, "CurrencyCode"

    EnumerateDocumentProperties projectFile, "Built-in document properties", "BuiltinDocumentProperties"
    EnumerateDocumentProperties projectFile, "Custom document properties", "CustomDocumentProperties"
End Sub

Sub EchoNamedProjectProperty(ByVal projectFile, ByVal propertyName)
    Dim value
    value = GetProjectPropertyValue(projectFile, propertyName)
    WScript.Echo propertyName & ": " & value
End Sub

Function GetProjectPropertyValue(ByVal projectFile, ByVal propertyName)
    On Error Resume Next
    Select Case propertyName
        Case "Name"
            GetProjectPropertyValue = projectFile.Name
        Case "FullName"
            GetProjectPropertyValue = projectFile.FullName
        Case "Path"
            GetProjectPropertyValue = projectFile.Path
        Case "Title"
            GetProjectPropertyValue = projectFile.Title
        Case "Subject"
            GetProjectPropertyValue = projectFile.Subject
        Case "Author"
            GetProjectPropertyValue = projectFile.Author
        Case "Manager"
            GetProjectPropertyValue = projectFile.Manager
        Case "Company"
            GetProjectPropertyValue = projectFile.Company
        Case "Comments"
            GetProjectPropertyValue = projectFile.Comments
        Case "CreationDate"
            GetProjectPropertyValue = projectFile.CreationDate
        Case "LastSaveDate"
            GetProjectPropertyValue = projectFile.LastSaveDate
        Case "Start"
            GetProjectPropertyValue = projectFile.Start
        Case "Finish"
            GetProjectPropertyValue = projectFile.Finish
        Case "StatusDate"
            GetProjectPropertyValue = projectFile.StatusDate
        Case "CurrentDate"
            GetProjectPropertyValue = projectFile.CurrentDate
        Case "Calendar"
            GetProjectPropertyValue = projectFile.Calendar
        Case "CurrencySymbol"
            GetProjectPropertyValue = projectFile.CurrencySymbol
        Case "CurrencyCode"
            GetProjectPropertyValue = projectFile.CurrencyCode
        Case Else
            GetProjectPropertyValue = ""
    End Select

    If Err.Number <> 0 Then
        GetProjectPropertyValue = "<unavailable: " & Err.Description & ">"
        Err.Clear
    ElseIf IsEmpty(GetProjectPropertyValue) Or IsNull(GetProjectPropertyValue) Then
        GetProjectPropertyValue = ""
    End If
    On Error GoTo 0
End Function

Sub EnumerateDocumentProperties(ByVal projectFile, ByVal heading, ByVal collectionName)
    Dim properties
    Dim propertyItem
    Dim propertyValue

    WScript.Echo ""
    WScript.Echo heading
    WScript.Echo String(Len(heading), "-")

    On Error Resume Next
    Select Case collectionName
        Case "BuiltinDocumentProperties"
            Set properties = projectFile.BuiltinDocumentProperties
        Case "CustomDocumentProperties"
            Set properties = projectFile.CustomDocumentProperties
    End Select

    If Err.Number <> 0 Or properties Is Nothing Then
        WScript.Echo "<unavailable: " & Err.Description & ">"
        Err.Clear
        On Error GoTo 0
        Exit Sub
    End If
    On Error GoTo 0

    For Each propertyItem In properties
        propertyValue = GetDocumentPropertyValue(propertyItem)
        WScript.Echo propertyItem.Name & ": " & propertyValue
    Next
End Sub

Function GetDocumentPropertyValue(ByVal propertyItem)
    On Error Resume Next
    GetDocumentPropertyValue = propertyItem.Value
    If Err.Number <> 0 Then
        GetDocumentPropertyValue = "<unavailable: " & Err.Description & ">"
        Err.Clear
    ElseIf IsEmpty(GetDocumentPropertyValue) Or IsNull(GetDocumentPropertyValue) Then
        GetDocumentPropertyValue = ""
    End If
    On Error GoTo 0
End Function