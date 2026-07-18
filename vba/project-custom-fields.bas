Attribute VB_Name = "main"
Option Explicit

Const DQUOTE As String = """"
Const DELIMITER As String = ","

Sub Main()

    Call ListPropertiesWithOffice16Library
    Call ListEveryCustomFieldTypeDynamically

End Sub

Sub ListPropertiesWithOffice16Library()
    ' Requires explicit reference to: Microsoft Office 16.0 Object Library
    Dim docProps As Office.DocumentProperties
    Dim docProp As Office.DocumentProperty
    Dim i As Long
    Dim customAlias As String
    Dim nativeName As String
    Dim fieldValue As String
    
    Dim fieldId As Long
    Dim fieldType As String
    
    Debug.Print "=== CUSTOM DOCUMENT PROPERTIES ==="
    Set docProps = ActiveProject.CustomDocumentProperties
    
    ' 1. Loop through File > Info > Advanced Properties > Custom
    Debug.Print "Scope" & DELIMITER & "Type" & DELIMITER & "Name" & DELIMITER & DQUOTE & "Custom Alias" & DQUOTE & DELIMITER & DQUOTE & "Value" & DQUOTE
    For Each docProp In docProps
        On Error Resume Next ' Bypasses custom links that are empty
        
        customAlias = ""
        fieldId = 0
        fieldType = ""
        nativeName = ""
        
        customAlias = docProp.Name
        fieldId = FieldNameToFieldConstant(customAlias, pjProject)
        fieldType = GetFieldTypeFromID(fieldId)
        nativeName = FieldConstantToFieldName(fieldId)
        
        'Debug.Print "Doc Property: " & docProp.Name & " = " & docProp.Value
        Debug.Print "Project" & DELIMITER & fieldType & DELIMITER & nativeName & DELIMITER & DQUOTE & customAlias & DQUOTE & DELIMITER & DQUOTE & docProp.Value & DQUOTE
        On Error GoTo 0
    Next docProp
    
    
End Sub

Sub ListEveryCustomFieldTypeDynamically()
    ' Requires explicit reference to: Microsoft Office 16.0 Object Library
   
    Dim fieldId As Long
    
    Debug.Print "Scope" & DELIMITER & "Type" & DELIMITER & "Name" & DELIMITER & DQUOTE & "Custom Alias" & DQUOTE
    
    ' Loop through the entire valid internal custom field ID enumeration blocks per type
    
    ' Task custom attributes
    For fieldId = 188743680 To 188745000 '188743680 To 188744150
        ListField (fieldId)
    Next fieldId
    
    ' Resource resource attributes
    For fieldId = 205520986 To 205522500 ' 205521003 To 205521226
        ListField (fieldId)
    Next fieldId
End Sub

Sub ListField(ByVal fieldId As Long)
    Dim customAlias As String
    Dim nativeName As String
    Dim fieldScope As String
    Dim fieldType As String
    
    customAlias = ""
    On Error Resume Next
    customAlias = CustomFieldGetName(fieldId)
    On Error GoTo 0
    
    ' If an alias exists, this structural slot is actively customized
    If customAlias <> "" Then
        nativeName = FieldConstantToFieldName(fieldId)
        
        ' Determine Scope (Task vs Resource)
        fieldScope = GetFieldScopeFromID(fieldId)
        
        ' Determine Data Type dynamically by parsing the string token
        fieldType = GetFieldTypeFromID(fieldId)
        
        Debug.Print fieldScope & DELIMITER & fieldType & DELIMITER & nativeName & DELIMITER & DQUOTE & customAlias & DQUOTE
    End If

End Sub

Public Function GetFieldScopeFromID(ByVal fieldId As Long) As String
    ' Evaluates the numeric FieldID block based on MS Project core architecture
    Select Case fieldId
        
        ' -------------------------------------------------------------
        ' 1. ENTERPRISE PROJECT FIELDS
        ' -------------------------------------------------------------
        Case 190873600 To 190875000
            GetFieldScopeFromID = "Enterprise Project"
            
        ' -------------------------------------------------------------
        ' 2. TASK SCOPE FIELDS
        ' -------------------------------------------------------------
        ' Covers Native Task fields (1 to ~30,000)
        ' Covers Local Task Custom fields (e.g., pjTaskText1 base 188,743,680)
        ' Covers Enterprise Task Custom fields (188,744,192+)
        Case 1 To 50000, 188743680 To 188745000
            GetFieldScopeFromID = "Task"
            
        ' -------------------------------------------------------------
        ' 3. RESOURCE SCOPE FIELDS
        ' -------------------------------------------------------------
        ' Covers Native Resource fields (51400+)
        ' Covers Local Resource Custom fields (e.g., pjResourceText1 base 205,520,986)
        ' Covers Enterprise Resource Custom fields (205,521,408+)
        Case 51400 To 60000, 205520986 To 205522500
            GetFieldScopeFromID = "Resource"
            
        ' -------------------------------------------------------------
        ' 4. FALLBACK / NOT FOUND
        ' -------------------------------------------------------------
        Case Else
            GetFieldScopeFromID = "Unknown"
            
    End Select
End Function

Public Function GetFieldTypeFromID(ByVal fieldId As Long) As String
    Dim nativeName As String
    
    ' 1. Pull the official baseline name from the MS Project engine
    On Error Resume Next
    nativeName = FieldConstantToFieldName(fieldId)
    On Error GoTo 0
    
    ' If the ID is invalid, exit early
    If nativeName = "" Then
        GetFieldTypeFromID = ""
        Exit Function
    End If
    
    ' 2. Clean up the string to identify the underlying data type
    ' This strips away "Task", "Resource", and any trailing numbers (e.g. "TaskText14" -> "Text")
    Dim cleanType As String
    cleanType = Replace(nativeName, "Task", "")
    cleanType = Replace(cleanType, "Resource", "")
    
    Dim i As Integer
    For i = 0 To 9
        cleanType = Replace(cleanType, CStr(i), "")
    Next i
    
    ' 3. Validate against standard MS Project Custom Field structural names
    Select Case cleanType
        Case "Text", "Date", "Cost", "Number", "Duration", "Flag", "Start", "Finish", "OutlineCode"
            ' Normalise layout formatting for standard outputs
            If cleanType = "OutlineCode" Then
                GetFieldTypeFromID = "Outline Code"
            Else
                GetFieldTypeFromID = cleanType
            End If
        Case Else
            GetFieldTypeFromID = "Standard/Unknown"
    End Select
End Function

