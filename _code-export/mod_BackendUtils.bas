Attribute VB_Name = "mod_BackendUtils"
Option Explicit


Public Function Model_Reset()
On Error Resume Next

    If MsgBox("Ok to FACTORY RESET this model?", vbExclamation + vbOKCancel, "Drop Model") <> vbOK Then Exit Function
    
    Dim ts As String: ts = "#" & Now & "#"
    
    With CurrentDb
    
        .Execute "DELETE * FROM app_Filestage"
        .Execute "DELETE * FROM app_Kvs"
        .Execute "DELETE * FROM app_Kvs_old"
        .Execute "DELETE * FROM app_Mailstage"
        .Execute "DELETE * FROM app_Filestage"
    
        .Execute "DELETE * FROM tbl__Comments"
        .Execute "DELETE * FROM tbl__Files"
        .Execute "DELETE * FROM tbl_Definitions_old"
        .Execute "DELETE * FROM tbl_Emails_old"
        .Execute "DELETE * FROM tbl_Inputs_old"
        .Execute "DELETE * FROM tbl_Memoranda_old"
        .Execute "DELETE * FROM tbl_Model_old;"
        .Execute "DELETE * FROM tbl_Organisations_old"
        .Execute "DELETE * FROM tbl_Outputs_old"
        .Execute "DELETE * FROM tbl_People_old"
        .Execute "DELETE * FROM tbl_Schemes_old"
        .Execute "DELETE * FROM tbl_WordPad_old"
        
        .Execute "DELETE * FROM tbl_Emails"
        .Execute "DELETE * FROM tbl_Inputs"
        .Execute "DELETE * FROM tbl_Memoranda"
        .Execute "DELETE * FROM tbl_Organisations"
        .Execute "DELETE * FROM tbl_Outputs"
        .Execute "DELETE * FROM tbl_People"
        .Execute "DELETE * FROM tbl_Schemes"
        .Execute "DELETE * FROM tbl_WordPad"

        .Execute "UPDATE tbl_Definitions SET sysStartTime = " & ts
        Definition_Verify "class"
        Definition_Verify "loc"
        Definition_Verify "orig"
        Definition_Verify "project"
        Definition_Verify "rev"
        Definition_Verify "role"
        Definition_Verify "scheme"
        Definition_Verify "sec"
        Definition_Verify "status"
        Definition_Verify "sys"
        Definition_Verify "type"
        Definition_Verify "ws"
        
        .Execute "UPDATE tbl_Model SET sysStartTime = " & ts
        .Execute "INSERT INTO tbl_Model (k,SysStartTime,v,SysStartPerson) VALUES ('is-editable'," & ts & ",'true',1)"
        .Execute "INSERT INTO tbl_Model (k,SysStartTime,v,SysStartPerson) VALUES ('ref-padding'," & ts & ",4,1)"
        .Execute "INSERT INTO tbl_Model (k,SysStartTime,v,SysStartPerson) VALUES ('rev-padding'," & ts & ",2,1)"
        .Execute "INSERT INTO tbl_Model (k,SysStartTime,v,SysStartPerson) VALUES ('wp-S'," & ts & ",'Empty Template',1)"
        .Execute "DELETE * FROM tbl_Model WHERE k = 'date'"
        .Execute "DELETE * FROM tbl_Model WHERE k = 'id'"
        .Execute "DELETE * FROM tbl_Model WHERE k = 'name'"
        .Execute "DELETE * FROM tbl_Model WHERE k = 'valid-hash'"
        
    End With
    
    DoCmd.DeleteObject acTable, "tbl_Definitions_cache"
    
End Function

Public Function History_Reset()
On Error Resume Next

    If MsgBox("Ok to CLEAR ALL HISTORY for this model?", vbExclamation + vbOKCancel, "Drop Model History") <> vbOK Then Exit Function
    
    With CurrentDb

        .Execute "DELETE * FROM app_Filestage"
        .Execute "DELETE * FROM app_Kvs"
        .Execute "DELETE * FROM app_Kvs_old"
        .Execute "DELETE * FROM app_Mailstage"
        .Execute "DELETE * FROM app_Filestage"
    
        .Execute "DELETE * FROM tbl_Definitions_old"
        .Execute "DELETE * FROM tbl_Emails_old;"
        .Execute "DELETE * FROM tbl_Inputs_old;"
        .Execute "DELETE * FROM tbl_Memoranda_old;"
        .Execute "DELETE * FROM tbl_Model_old;"
        .Execute "DELETE * FROM tbl_Organisations_old;"
        .Execute "DELETE * FROM tbl_Outputs_old;"
        .Execute "DELETE * FROM tbl_People_old;"
        .Execute "DELETE * FROM tbl_Schemes_old;"
        .Execute "DELETE * FROM tbl_WordPad_old;"
        
    End With
    
    DoCmd.DeleteObject acTable, "tbl_Definitions_cache"
    
End Function

Private Sub Definition_Verify(codeName As String)
    
    Dim count As Long: count = DCount("code", "tbl_Definitions", "field='" & codeName & "Code'")
    If count <= 0 Then MsgBox "No " & codeName & "Codes found." & vbCrLf & vbCrLf & "Note: The application cannot function without a complete / consistent Definitions table. The model manager should review & update these prior to work commencing.", vbExclamation, "Error: Incomplete Definitions"
    
End Sub
