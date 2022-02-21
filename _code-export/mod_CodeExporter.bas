Attribute VB_Name = "mod_CodeExporter"
Const VB_MODULE = 1
Const VB_CLASS = 2
Const VB_FORM = 100
Const EXT_MODULE = ".bas"
Const EXT_CLASS = ".cls"
Const EXT_FORM = ".frm"
Const CODE_FLD = "_code-export"

Sub ExportAllCode()

Dim fileName As String
Dim exportPath As String
Dim ext As String
Dim FSO As Object

Set FSO = CreateObject("Scripting.FileSystemObject")
' Set export path and ensure its existence
exportPath = CurrentProject.Path & "\" & CODE_FLD
If Not FSO.FolderExists(exportPath) Then
    MkDir exportPath
End If

' The loop over all modules/classes/forms
For Each C In Application.VBE.VBProjects(1).VBComponents
    ' Get the filename extension from type
    ext = vbExtFromType(C.Type)
    If ext <> "" Then
        fileName = C.name & ext
        Debug.Print "Exporting " & C.name & " to file " & fileName
        ' THE export
        C.Export exportPath & "\" & fileName
    Else
        Debug.Print "Unknown VBComponent type: " & C.Type
    End If
Next C

End Sub

' Helper function that translates VBComponent types into file extensions
' Returns an empty string for unknown types
Function vbExtFromType(ByVal ctype As Integer) As String
    Select Case ctype
        Case VB_MODULE
            vbExtFromType = EXT_MODULE
        Case VB_CLASS
            vbExtFromType = EXT_CLASS
        Case VB_FORM
            vbExtFromType = EXT_FORM
    End Select
End Function
