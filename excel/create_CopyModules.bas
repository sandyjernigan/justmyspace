Attribute VB_Name = "create_CopyModules"
'Call CopyAllModules(TargetFile)
    ' ^^ copies all modules in this workbook to a new workbook

'Call CopyModule(strModuleName, TargetFile)
    ' ^^ copies a module from one workbook to another

Sub CopyAllModules() 'TargetFile As String

    Dim modArray As Variant
    Dim thisMod As String
    
    modArray = Array("a_Variables", _
        "create_CopyModules", _
        "Shared_AgentFunctions", _
        "Shared_DateFunctions", _
        "Shared_Functions", _
        "Shared_Styling")
    
    For i = 0 To UBound(modArray)
        thisMod = modArray(i)
        
        If (Len(thisMod) > 0) Then
        
            Call CopyModule(thisMod, "Book1.xlsx")
            
        End If
    Next

End Sub

Sub CopyModule(strModuleName As String, TargetFile As String)
' copies a module from one workbook to another

    Dim strFolder As String, strTempFile As String
    
    'Set Work Books
    Set SourceWB = ThisWorkbook
    'Set TargetWB = Workbooks("Book1.xlsx")
    
    'Get Source Path
    strFolder = SourceWB.path
    If Len(strFolder) = 0 Then strFolder = CurDir
    strFolder = strFolder & "\Macros\"
    
    'Set Temporary File Name
    strTempFile = strFolder & strModuleName & ".bas"
    
    'On Error Resume Next
    
    'Export and Import Module
    SourceWB.VBProject.VBComponents(strModuleName).Export strTempFile
    'TargetWB.VBProject.VBComponents.Import strTempFile
    
    'Delete Temporary File
    'Kill strTempFile
    
    On Error GoTo 0
End Sub




