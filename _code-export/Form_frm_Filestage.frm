VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frm_Filestage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const prefix As String = "N"
Private Const suffix As String = "Inputs"
Public w As Long
Public h As Long


Private Sub btn_ApplyAll_Click()
    If MsgBox("Click OK to apply ALL PRESETS to the selected items?", vbOKCancel, "Apply ALL") <> vbOK Then Exit Sub
    Me.Dirty = False
    sub_Browser.Form.SelectedRecords_ApplyPresets
    sub_Browser.Requery
End Sub

Private Sub btn_Commit_Click()
    Dim rst As Recordset: Set rst = dbLocal.OpenRecordset("SELECT target, id from app_Filestage WHERE target IS NOT NULL AND NOT IsNumeric(target) ORDER BY target;", dbOpenForwardOnly)
    Dim trg As String
    Dim d As Dictionary: Set d = New Dictionary
    Do Until rst.EOF
        trg = rst.Fields("target").value
        If Not d.Exists(trg) Then d.Add trg, vbNullString
        d(trg) = CsvAdd_numSort(d(trg), rst.Fields("id").value, True)
        rst.MoveNext
    Loop
    rst.Close
    Set rst = Nothing
    
    If d.count = 0 Then
        MsgBox "No commitable / green files were found.", vbInformation, "Commit Inputs - Failed"
        Exit Sub
    End If
    If MsgBox("OK to commit the green / relevant files into the model Inputs register?", vbOKCancel, "Commit Inputs") <> vbOK Then Exit Sub
    Dim vKey As Variant
    For Each vKey In d.Keys()
        If g_Model.Input_Commit(d(vKey)) > 0 Then dbLocal.Execute "DELETE * FROM app_Filestage WHERE ID IN(" & d(vKey) & ");"
    Next
    sub_Browser.Requery
On Error Resume Next
    Form_frm_Inputs.sub_Browser.Requery
    Set d = Nothing
End Sub

Private Sub btn_Merge_Click()
    Dim rst As Recordset: Set rst = dbLocal.OpenRecordset("SELECT target, id from app_Filestage WHERE target IS NOT NULL AND IsNumeric(target) ORDER BY target;", dbOpenForwardOnly)
    Dim trg As String
    Dim d As Dictionary: Set d = New Dictionary
    Do Until rst.EOF
        trg = rst.Fields("target").value
        If Not d.Exists(trg) Then d.Add trg, vbNullString
        If d(trg) = vbNullString Then
            d(trg) = rst.Fields("id").value
        Else
            d(trg) = d(trg) & "," & rst.Fields("id").value
        End If
        rst.MoveNext
    Loop
    rst.Close
    Set rst = Nothing
        
    If d.count = 0 Then
        MsgBox "No mergeable / blue files were found.", vbInformation, "Merge Inputs - Failed"
        Exit Sub
    End If
    
    If MsgBox("OK to merge the blue / relevant files into the model Inputs register?", vbOKCancel, "Merge Inputs") <> vbOK Then Exit Sub
    Dim vKey As Variant
    For Each vKey In d.Keys()
        If g_Model.Input_Merge(CLng(vKey), d(vKey)) > 0 Then dbLocal.Execute "DELETE * FROM app_Filestage WHERE ID IN(" & d(vKey) & ");"
    Next
    sub_Browser.Requery
    
On Error Resume Next
    Form_frm_Inputs.sub_Browser.Requery
    Set d = Nothing
End Sub

Private Sub btn_Refresh_Click()
    With sub_Browser
        .Form.CheckMatches
        MsgBox "All potential matches / hits have been updated.", vbInformation, "Refresh Complete!"
        .Requery
    End With
End Sub

Private Sub chk_AllNone_DblClick(Cancel As Integer)
    With sub_Browser
        If .Form.VisibleRecords_Count > 0 Then
            Select Case chk_AllNone
                Case True:
                    chk_AllNone = False
                Case False:
                    chk_AllNone = True
                Case Else:
                    If MsgBox("Click OK to clear the current selection?", vbOKCancel, "Clear Selection") <> vbOK Then Exit Sub
                    chk_AllNone = False
            End Select
            .Form.VisibleRecords_Selected chk_AllNone
            .Requery
        End If
    End With
End Sub

Private Sub classCode_DblClick(Cancel As Integer)
    DoCmd.OpenForm "fdlg_Classifier", , , , , , Me.name & ";classCode"
End Sub

Private Sub effectiveOrg_GotFocus()
    effectiveOrg.Requery
End Sub

Private Sub effectiveOrg_NotInList(newData As String, Response As Integer)
    Response = Organisation_CreateNotInList(effectiveOrg, newData)
End Sub

Private Sub Form_Load()
    
    ' Set_FormPermissions
    Set_FormIcon Me, LCase$(suffix)
    Set_FormSize Me
    
    
    Dim bln_IsReadOnly As Boolean: bln_IsReadOnly = g_Model.Db_isReadOnly
    Dim bln_IsFsConnected As Boolean: bln_IsFsConnected = g_Model.Fs_isConnected
    Dim ctl As Control
    For Each ctl In Me.Controls
        Select Case True
            Case InStr(ctl.Tag, "rw") > 0: ctl.Enabled = Not bln_IsReadOnly
            Case InStr(ctl.Tag, "fs") > 0: ctl.Enabled = bln_IsFsConnected
        End Select
    Next ctl
    sub_Browser.Form.AllowEdits = Not bln_IsReadOnly
    
    ' Form_Init
    FormHeader.BackColor = Pastel_1.Green
    Detail.BackColor = Pastel_1.Green
    FormFooter.BackColor = Pastel_1.Green
    Me.Caption = "Filestage"
    
    Dim li As ListItem
    With lv_FileDrop
        .ColumnHeaders.Clear
        .ColumnHeaders.Add , , "txt", 100        ', 2000
        .ListItems.Clear
        'Set li = .ListItems.Add(, , "  Drop FILES here")
    End With
    With lv_FolderDrop
        .ColumnHeaders.Clear
        .ColumnHeaders.Add , , "txt", 100        ', 2000
        .ListItems.Clear
        'Set li = .ListItems.Add(, , "  Drop FOLDERS here")
    End With
    
    With sub_Browser.Form
        Refresh_Controls .VisibleRecords_Count > 0
        .CheckMatches
    End With
    
    Refresh_Captions
    
End Sub

Private Sub Form_Open(Cancel As Integer)
    If Not g_Model.Hasher_Validate Then
        MsgBox "hasher.exe did not return the expected hash value for the system validation case. Therefore to avoid creating inconsistencies in this model, filestage access is currently denied.", vbExclamation, "Invalid Plugin: hasher.exe"
        DoCmd.Close acForm, Me.name, acSaveNo
    End If
End Sub

Private Sub Form_Resize()
    Limit_FormSize Me
End Sub

Private Sub lbl_classCode_DblClick(Cancel As Integer)
    If Not classCode.Enabled Then Exit Sub
    Me.Dirty = False
    With sub_Browser
        .Form.SelectedRecords_SetString classCode.name, Nz(classCode, vbNullString)
        .Requery
    End With
End Sub

Private Sub lbl_Clear_DblClick(Cancel As Integer)
    effectiveOrg = Null
    TypeCode = vbNullString
    secCode = vbNullString
    classCode = vbNullString
    statusCode = vbNullString
End Sub

Private Sub lbl_effectiveOrg_DblClick(Cancel As Integer)
    If Not effectiveOrg.Enabled Then Exit Sub
    Me.Dirty = False
    With sub_Browser
        .Form.SelectedRecords_SetLong effectiveOrg.name, Nz(effectiveOrg, 0)
        .Requery
    End With
End Sub

Private Sub lbl_Folders_Click()
    Execute g_App.App_Assets & "tree-visualiser.jar"
End Sub

Private Sub lbl_Inputs_Click()
    DoCmd.OpenForm "frm_Inputs"
End Sub

Private Sub lbl_secCode_DblClick(Cancel As Integer)
    If Not secCode.Enabled Then Exit Sub
    Me.Dirty = False
    With sub_Browser
        .Form.SelectedRecords_SetString secCode.name, Nz(secCode, vbNullString)
        .Requery
    End With
End Sub

Private Sub lbl_statusCode_DblClick(Cancel As Integer)
    If Not statusCode.Enabled Then Exit Sub
    Me.Dirty = False
    With sub_Browser
        .Form.SelectedRecords_SetString statusCode.name, Nz(statusCode, vbNullString)
        .Requery
    End With
End Sub

Private Sub lbl_target_DblClick(Cancel As Integer)
    If Not target.Enabled Then Exit Sub
    Me.Dirty = False
    With sub_Browser
        .Form.SelectedRecords_SetString target.name, Nz(target, vbNullString)
        .Requery
    End With
    Me.target = Null
End Sub

Private Sub lbl_typeCode_DblClick(Cancel As Integer)
    If Not TypeCode.Enabled Then Exit Sub
    Me.Dirty = False
    With sub_Browser
        .Form.SelectedRecords_SetString TypeCode.name, Nz(TypeCode, vbNullString)
        .Requery
    End With
End Sub

Private Sub lv_FileDrop_OLEDragDrop(Data As Object, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo MyExit

    Dim col As New Collection
    Dim hasRejects As Boolean
    Dim i As Long
    For i = 1 To Data.files.count
        If ValidFile(Data.files(i)) Then
            col.Add Data.files(i)
        Else
            hasRejects = True
        End If
    Next i
    If hasRejects Then MsgBox Filestage_FileDropWarning1, vbInformation & vbOKOnly, "Invalid Files Detected"
    
    If col.count > 0 Then
        Dim status As String
        Dim vItem As Variant
        For Each vItem In col
            status = "Error:   "
            If g_Model.File_Stage(CStr(vItem)) Then status = "Staged:   "
            g_App.StatusBar_Set status & CStr(vItem)
            sub_Browser.Requery
        Next
    Else
        MsgBox "No valid Files were detected!", vbInformation, "Filedrop"
    End If
    
    g_App.StatusBar_Clear
    Exit Sub

MyExit:
    MsgBox "The dropped content was not recognised.", vbExclamation, "Invalid Filedrop"
End Sub

Private Sub lv_FolderDrop_OLEDragDrop(Data As Object, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo MyExit

    Dim col As New Collection
    
    Dim i As Long
    For i = 1 To Data.files.count
        If Folder_Exists(Data.files(i)) Then RecursiveDir Data.files(i), col
    Next i
    
    Dim hasRejects As Boolean
    For i = col.count To 1 Step -1
        If Not ValidFile(col(i)) Then
            col.Remove i
            hasRejects = True
        End If
    Next i
    If hasRejects Then MsgBox Filestage_FileDropWarning1, vbInformation & vbOKOnly, "Invalid Files Detected"
    
    If col.count > 0 Then
        Dim status As String
        Dim vItem As Variant
        For Each vItem In col
            status = "Error:   "
            If g_Model.File_Stage(CStr(vItem)) Then status = "Staged:   "
            g_App.StatusBar_Set status & CStr(vItem)
            sub_Browser.Requery
        Next
    Else
        MsgBox "No valid Files were detected!", vbInformation, "Folderdrop"
    End If
    
    g_App.StatusBar_Clear
    
    Exit Sub

MyExit:
    MsgBox "The dropped content was not recognised.", vbExclamation, "Invalid Folderdrop"
End Sub

Private Sub secCode_DblClick(Cancel As Integer)
    DoCmd.OpenForm "fdlg_Classifier", , , , , , Me.name & ";secCode"
End Sub

Private Sub statusCode_DblClick(Cancel As Integer)
    DoCmd.OpenForm "fdlg_Classifier", , , , , , Me.name & ";statusCode"
End Sub

Private Sub txt_url_Change()
    Dim col As New Collection
    Dim strArr() As String: strArr = Split(txt_url.Text, vbCrLf)
    Dim strUrl As String
    
    Dim vItem As Variant
    For Each vItem In strArr
        strUrl = ValidUrl(CStr(vItem))
        If strUrl <> vbNullString Then col.Add strUrl
    Next
      
    If col.count > 0 Then
        Dim status As String
        For Each vItem In col
            status = "Error:   "
            If g_Model.Url_Stage(CStr(vItem)) Then status = "Staged:   "
            g_App.StatusBar_Set status & CStr(vItem)
        Next
    Else
        MsgBox "No valid URLs were detected!", vbInformation, "Paste URL"
    End If
    
    MsgBox "Complete!", vbInformation, "Stage Urls"
    txt_url = vbNullString
    sub_Browser.Requery
    g_App.StatusBar_Clear
    
End Sub

Private Sub txt_url_Click()
    If Len(txt_url & "") = 0 Then Exit Sub
    txt_url.SelStart = 0
    txt_url.SelLength = Len(txt_url)
End Sub

Private Sub typeCode_DblClick(Cancel As Integer)
    DoCmd.OpenForm "fdlg_Classifier", , , , , , Me.name & ";typeCode"
End Sub

' ------- '
' PRIVATE '
' ------- '

' ------ '
' PUBLIC '
' ------ '

Public Sub Refresh_Controls(Optional bln_hasRecords As Boolean = False)

    If g_Model.Db_isReadOnly Then Exit Sub

    chk_AllNone.Enabled = bln_hasRecords
    TypeCode.Enabled = bln_hasRecords
    secCode.Enabled = bln_hasRecords
    classCode.Enabled = bln_hasRecords
    statusCode.Enabled = bln_hasRecords
    target.Enabled = bln_hasRecords
    effectiveOrg.Enabled = bln_hasRecords
    btn_ApplyAll.Enabled = bln_hasRecords
    btn_Refresh.Enabled = bln_hasRecords
    
    btn_Commit.Enabled = g_Model.Fs_isConnected And bln_hasRecords
    btn_Merge.Enabled = btn_Commit.Enabled
    
    txt_url = vbNullString
    
    If bln_hasRecords Then
        With sub_Browser.Form
            Dim bln_Checked As Boolean: bln_Checked = .VisibleRecords_Checked
            Dim bln_Unchecked As Boolean: bln_Unchecked = .VisibleRecords_UnChecked
        End With
        If bln_Checked And bln_Unchecked Then
            chk_AllNone = Null
        Else
            chk_AllNone = bln_Checked
        End If
    Else
        chk_AllNone = False
    End If
    
End Sub

Public Sub Refresh_Captions()
On Error Resume Next
    Dim fullpath As String: fullpath = Nz(sub_Browser.Form!fullpath, vbNullString)
    If fullpath = vbNullString Then
        Caption = "Filestage   ~   No selection"
    Else
        Caption = "Filestage   ~   " & fullpath
    End If
End Sub

Public Function ValidFile(strPath As String) As Boolean
    ValidFile = False
    With New FileSystemObject
        If Not .FileExists(strPath) Then Exit Function
        Select Case .GetExtensionName(strPath)
            Case "eml": Exit Function
            Case "msg": Exit Function
            Case "url": Exit Function
            Case "lnk": Exit Function
        End Select
    End With
    ValidFile = True
End Function
