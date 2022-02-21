VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frm_Mailstage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const prefix As String = "E"
Private Const suffix As String = "Emails"
Public w As Long
Public h As Long


Private Sub btn_ApplyAll_Click()
    If MsgBox("Click OK to apply ALL PRESETS to the selected items?", vbOKCancel, "Apply ALL") <> vbOK Then Exit Sub
    If Me.Dirty Then Me.Dirty = False
    sub_Browser.Form.SelectedRecords_ApplyPresets
    sub_Browser.Requery
End Sub

Private Sub btn_Commit_Click()

    Dim rst As Recordset: Set rst = sub_Browser.Form.VisibleRecords
    If (rst.EOF And rst.BOF) Then Exit Sub
    
    If Not sub_Browser.Form.VisibleRecords_Checked Then
        MsgBox "There are no emails selected for committing.", vbExclamation, "Commit Emails - Failed"
        Exit Sub
    End If

    If MsgBox("OK to commit the selected emails into the Emails register?", vbQuestion + vbOKCancel, "Commit Emails") <> vbOK Then Exit Sub
    
    rst.MoveFirst
    Do Until rst.EOF = True
        If Not rst!selected Then GoTo NextOne
        If Nz(rst!mcount, 0) = 0 Then GoTo Commit
        If MsgBox("...\" & rst!fname & "." & rst!ftype & " clashes with " & rst!mcount & " other emails." & vbCrLf & vbCrLf & "Are you sure you want to commit this potential duplicate?", vbExclamation + vbYesNo, "Commit Duplicate?") <> vbYes Then GoTo NextOne
Commit:
        If g_Model.Email_Create(rst!ID) > 0 Then dbLocal.Execute "DELETE * FROM app_Mailstage WHERE ID = " & rst!ID & ";"
NextOne:
        rst.MoveNext
    Loop
    
    rst.Close
    Set rst = Nothing
    
On Error Resume Next
    sub_Browser.Requery
    Form_frm_Emails.sub_Browser.Requery
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
    Set_FormSize Me
    Set_FormIcon Me, LCase$(suffix)
    
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
    FormHeader.BackColor = Pastel_1.Grape
    Detail.BackColor = Pastel_1.Grape
    FormFooter.BackColor = Pastel_1.Grape
    Me.Caption = "Mailstage"
    
    Dim li As ListItem
    With lv_FileDrop
        .ColumnHeaders.Clear
        .ColumnHeaders.Add , , "txt", 100        ', 2000
        .ListItems.Clear
        'Set li = .ListItems.Add(, , "  Drop FILES here")
    End With
    
    With sub_Browser.Form
        Refresh_Controls .VisibleRecords_Count > 0
        .CheckMatches
    End With
    
    Refresh_Captions
    
End Sub

Private Sub Form_Open(Cancel As Integer)
    Select Case True
        Case Not g_Model.Hasher_Validate: MsgBox "hasher.exe did not return the expected hash value for the system validation case. Therefore to avoid creating inconsistencies in this model, mailstage access is currently denied.", vbExclamation, "Invalid Plugin: hasher.exe"
        Case Not g_Model.MailParser_Validate: MsgBox "mail-parser.exe did not return the expected validation data. Therefore to avoid creating inconsistencies in this model, mailstage access is currently denied.", vbExclamation, "Invalid Plugin: mail-parser.exe"
        Case Else: Exit Sub
    End Select
    DoCmd.Close acForm, Me.name, acSaveNo
End Sub

Private Sub Form_Resize()
    Limit_FormSize Me
End Sub

Private Sub lbl_classCode_DblClick(Cancel As Integer)
    If Not classCode.Enabled Then Exit Sub
    If Me.Dirty Then Me.Dirty = False
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
    If Me.Dirty Then Me.Dirty = False
    With sub_Browser
        .Form.SelectedRecords_SetLong effectiveOrg.name, Nz(effectiveOrg, 0)
        .Requery
    End With
End Sub

Private Sub lbl_Emails_Click()
    DoCmd.OpenForm "frm_Emails"
End Sub

Private Sub lbl_secCode_DblClick(Cancel As Integer)
    If Not secCode.Enabled Then Exit Sub
    If Me.Dirty Then Me.Dirty = False
    With sub_Browser
        .Form.SelectedRecords_SetString secCode.name, Nz(secCode, vbNullString)
        .Requery
    End With
End Sub

Private Sub lbl_statusCode_DblClick(Cancel As Integer)
    If Not statusCode.Enabled Then Exit Sub
    If Me.Dirty Then Me.Dirty = False
    With sub_Browser
        .Form.SelectedRecords_SetString statusCode.name, Nz(statusCode, vbNullString)
        .Requery
    End With
End Sub

Private Sub lbl_typeCode_DblClick(Cancel As Integer)
    If Not TypeCode.Enabled Then Exit Sub
    If Me.Dirty Then Me.Dirty = False
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
        If ValidEmail(Data.files(i)) Then
            col.Add Data.files(i)
        Else
            Debug.Print "Invalid Email: " & Data.files(i)
            hasRejects = True
        End If
    Next i
    If hasRejects Then MsgBox Mailstage_FileDropWarning1, vbInformation & vbOKOnly, "Invalid Emails Detected"
    
    If col.count > 0 Then
        Dim vItem As Variant
        For Each vItem In col
            g_App.StatusBar_Set "Staging: " & CStr(vItem)
            If Not g_Model.Email_Stage(CStr(vItem)) Then Debug.Print "Failed to Stage: " & CStr(vItem)
            sub_Browser.Requery
        Next
    Else
        MsgBox "No valid Emails were detected!", vbInformation, "Maildrop"
    End If
    
    g_App.StatusBar_Clear
    'MsgBox "Complete!", vbInformation, "Stage Emails"
    
    Exit Sub

MyExit:
    MsgBox "The dropped content was not recognised.", vbExclamation, "Invalid Maildrop"
End Sub

Private Sub secCode_DblClick(Cancel As Integer)
    DoCmd.OpenForm "fdlg_Classifier", , , , , , Me.name & ";secCode"
End Sub

Private Sub statusCode_DblClick(Cancel As Integer)
    DoCmd.OpenForm "fdlg_Classifier", , , , , , Me.name & ";statusCode"
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
    effectiveOrg.Enabled = bln_hasRecords
    btn_ApplyAll.Enabled = bln_hasRecords
    btn_Refresh.Enabled = bln_hasRecords
    
    btn_Commit.Enabled = g_Model.Fs_isConnected And bln_hasRecords
    
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
    Dim subject As String: subject = Nz(sub_Browser.Form!subject, vbNullString)
    If subject = vbNullString Then
        Caption = "Mailstage   ~   No selection"
    Else
        Caption = "Mailstage   ~   " & subject
    End If
End Sub

Public Function ValidEmail(strPath As String) As Boolean
    ValidEmail = False
    With New FileSystemObject
        If Not .FileExists(strPath) Then Exit Function
        Select Case .GetExtensionName(strPath)
            Case "eml":
            Case "msg":
            Case Else: Exit Function
        End Select
    End With
    ValidEmail = True
End Function
