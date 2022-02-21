VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_fsub_Schemes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const prefix As String = "S"
Private Const suffix As String = "Schemes"
Private s_recordsForDeletion As String
Private ts_recordsForDeletion As Date


Private Sub accountable_GotFocus()
    accountable.Requery
End Sub

Private Sub accountable_NotInList(newData As String, Response As Integer)
    Response = Person_CreateNotInList(accountable, newData)
End Sub

Private Sub classCode_DblClick(Cancel As Integer)
    Manual_Classify Me
End Sub

Private Sub classCode_Enter()
    Auto_Classify Me
End Sub

Private Sub classCode_GotFocus()
    classCode.Requery
End Sub

Private Sub comment_DblClick(Cancel As Integer)
    CommentOn Me
End Sub

Private Sub Dir_Click()
    DirClick Me
End Sub

Private Sub dir_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = acRightButton And Nz(Me!dir, 0) > 0 Then
        Refresh_ContextMenu
        CommandBars("menu_" & suffix).ShowPopup
        DoCmd.CancelEvent
    End If
End Sub

Private Sub Form_AfterDelConfirm(status As Integer)
    If status <> acDeleteOK Then
        Dim r As Variant
        For Each r In Split(s_recordsForDeletion, ",")
            g_Model.Entity_Update "tbl_" & suffix, CLng(r), vbNullString, ts_recordsForDeletion
        Next r
    End If
    s_recordsForDeletion = vbNullString
    ts_recordsForDeletion = vbMaxDate
End Sub

Private Sub Form_AfterInsert()
    g_Model.Dir_Create prefix, Me!dir
End Sub

Private Sub Form_BeforeDelConfirm(Cancel As Integer, Response As Integer)
    Cancel = g_Model.Db_isReadOnly
    If Cancel Then Exit Sub
    s_recordsForDeletion = mod_StringUtils.TrimTrailingChr(s_recordsForDeletion, ",")
End Sub

Private Sub Form_BeforeInsert(Cancel As Integer)
    Cancel = g_Model.Db_isReadOnly
    If Cancel Then Exit Sub
    SetDefaults_Create
End Sub

Private Sub Form_BeforeUpdate(Cancel As Integer)
    Dim strMsg, strTitle As String: strTitle = "Update Denied - "
    Select Case True
        Case g_Model.Db_isReadOnly:
            strTitle = strTitle & "Read-only Mode"
            strMsg = "The model cannot be updated using FMMS viewer or whilst in read-only mode."
        Case Me.isFrozen:
            strTitle = strTitle & "Frozen Record"
            strMsg = "This record is currently frozen."
        Case Not g_Model.Entity_Archive("tbl_" & suffix, Me!dir):
            strTitle = strTitle & "Archiving Failed"
            strMsg = "The record couldn't be archived, so the update cannot proceed."
        Case Else: GoTo Execute
    End Select
error:
    MsgBox strMsg, vbExclamation, strTitle
    Me.Undo
    Cancel = True
    Exit Sub
Execute:
    SetDefaults_Update
End Sub

Private Sub Form_Current()
On Error Resume Next
    Form_frm_Schemes.Set_Criteria
    If CurrentProject.AllForms("frm_People").IsLoaded Then
        If Forms!frm_People!cmb_Main > 1 Then Form_frm_People.Refresh_RecordSource
    End If
    If CurrentProject.AllForms("frm_Organisations").IsLoaded Then
        If Forms!frm_Organisations!cmb_Main > 1 Then Form_frm_Organisations.Refresh_RecordSource
    End If
    If CurrentProject.AllForms("frm_Emails").IsLoaded Then
        If Forms!frm_Emails!cmb_Main > 1 Then Form_frm_Emails.Refresh_RecordSource
    End If
    If CurrentProject.AllForms("frm_Inputs").IsLoaded Then
        If Forms!frm_Inputs!cmb_Main > 1 Then Form_frm_Inputs.Refresh_RecordSource
    End If
    If CurrentProject.AllForms("frm_Memoranda").IsLoaded Then
        If Forms!frm_Memoranda!cmb_Main > 1 Then Form_frm_Memoranda.Refresh_RecordSource
    End If
    If CurrentProject.AllForms("frm_Outputs").IsLoaded Then
        If Forms!frm_Outputs!cmb_Main > 1 Then Form_frm_Outputs.Refresh_RecordSource
    End If
End Sub

Private Sub Form_Delete(Cancel As Integer)
    If ts_recordsForDeletion = vbMaxDate Then ts_recordsForDeletion = Now
    s_recordsForDeletion = s_recordsForDeletion & Me!dir & ","
    If g_Model.Entity_Archive("tbl_" & suffix, Me!dir) Then Exit Sub
abort:
    MsgBox "The record couldn't be archived, so the deletion cannot proceed.", vbExclamation, "Delete Denied - Archiving Failed"
    Cancel = True
End Sub

Private Sub Form_Load()
    ts_recordsForDeletion = vbMaxDate
End Sub

Private Sub Form_Open(Cancel As Integer)
    Query_Refresh "quni_" & suffix, std_Sql.Schemes_quni
    Me.RecordSource = "qry_" & suffix
End Sub

Private Sub indent_AfterUpdate()
    If Nz(Me.Title, vbNullString) <> vbNullString Then Me.Title = TreeIndent_Set(Me.Title, Me.indent)
End Sub

Private Sub indent_BeforeUpdate(Cancel As Integer)
    Cancel = IsNull(Me.indent) Or Me.indent > 2 Or Me.indent < 0
End Sub

Private Sub link_AfterUpdate()
    If link = vbNullString Then Exit Sub
    Dim URL As String: URL = mod_StringUtils.HyperlinkMidPart(link)
    link = "Open#" & URL & "##" & URL
End Sub

Private Sub Link_Click()
    LinkClick Me
End Sub

Private Sub locCode_DblClick(Cancel As Integer)
    Manual_Classify Me
End Sub

Private Sub locCode_Enter()
    Auto_Classify Me
End Sub

Private Sub locCode_GotFocus()
    locCode.Requery
End Sub

Private Sub origCode_DblClick(Cancel As Integer)
    Manual_Classify Me
End Sub

Private Sub origCode_Enter()
    Auto_Classify Me
End Sub

Private Sub origCode_GotFocus()
    origCode.Requery
End Sub

Private Sub responsible_GotFocus()
    responsible.Requery
End Sub

Private Sub responsible_NotInList(newData As String, Response As Integer)
    Response = Person_CreateNotInList(responsible, newData)
End Sub

Private Sub revCode_DblClick(Cancel As Integer)
    Manual_Classify Me
End Sub

Private Sub revCode_Enter()
    Auto_Classify Me
End Sub

Private Sub revCode_GotFocus()
    revCode.Requery
End Sub

Private Sub roleCode_DblClick(Cancel As Integer)
    Manual_Classify Me
End Sub

Private Sub roleCode_Enter()
    Auto_Classify Me
End Sub

Private Sub roleCode_GotFocus()
    roleCode.Requery
End Sub

Private Sub schemeCode_GotFocus()
    schemeCode.Requery
End Sub

Private Sub secCode_DblClick(Cancel As Integer)
    Manual_Classify Me
End Sub

Private Sub secCode_Enter()
    Auto_Classify Me
End Sub

Private Sub secCode_GotFocus()
    secCode.Requery
End Sub

Private Sub statusCode_DblClick(Cancel As Integer)
    Manual_Classify Me
End Sub

Private Sub statusCode_Enter()
    Auto_Classify Me
End Sub

Private Sub statusCode_GotFocus()
    statusCode.Requery
End Sub

Private Sub sysCode_DblClick(Cancel As Integer)
    Manual_Classify Me
End Sub

Private Sub sysCode_Enter()
    Auto_Classify Me
End Sub

Private Sub sysCode_GotFocus()
    sysCode.Requery
End Sub

Private Sub sysEndPerson_GotFocus()
    sysEndPerson.Requery
End Sub

Private Sub sysStartPerson_GotFocus()
    sysStartPerson.Requery
End Sub

Private Sub Title_AfterUpdate()
    Me.Title = TreeIndent_Set(Me.Title, Me.indent)
End Sub

Private Sub Title_BeforeUpdate(Cancel As Integer)
    Cancel = IsNull(Me.Title)
End Sub

Private Sub typeCode_DblClick(Cancel As Integer)
    Manual_Classify Me
End Sub

Private Sub typeCode_Enter()
    Auto_Classify Me
End Sub

Private Sub typeCode_GotFocus()
    TypeCode.Requery
End Sub

Private Sub wsCode_DblClick(Cancel As Integer)
    Manual_Classify Me
End Sub

Private Sub wsCode_Enter()
    Auto_Classify Me
End Sub

Private Sub wsCode_GotFocus()
    wsCode.Requery
End Sub

' ------- '
' PRIVATE '
' ------- '

Private Sub Refresh_ContextMenu()
    
    On Error Resume Next
        CommandBars("menu_" & suffix).Delete
    On Error GoTo 0
    
    Dim cbar As CommandBar: Set cbar = CommandBars.Add(name:="menu_" & suffix, Position:=msoBarPopup)
    
    With New cls_Permissor
        
        .Init Me.Form
        
        AddCustomButton(cbar, "Delivery Plan", "Scheme_WordPad", faceId_DeliveryPlan).BeginGroup = True
        
        If .Can_ReadFs Then
            AddCustomButton cbar, "Open (WIP)", "Scheme_OpenWIP", faceId_WorkInProgress
            AddCustomButton cbar, "Latest Snapshot", "Scheme_OpenLatestSnapshot", faceId_SnapshotLatest
            AddCustomButton cbar, "All Snapshots", "Scheme_OpenSnapshots", faceId_SnapshotAll
        End If
        
        AddCustomButton cbar, "Comments", "Scheme_Comments", faceId_Comments
        
        Dim sub_Edit As CommandBarControl: Set sub_Edit = cbar.Controls.Add(Type:=msoControlPopup)
        sub_Edit.Caption = "Edit"
        sub_Edit.Enabled = .Can_EditModel
        If .Can_TakeSnapshot Then AddCustomSubButton sub_Edit, "Take Snapshot", "Scheme_CreateSnapshot", faceId_SnapshotSave
        If .Can_Unfreeze Then AddCustomSubButton(sub_Edit, "Unfreeze", "Scheme_IsFrozen", faceId_Unfreeze).BeginGroup = True
        If .Can_Freeze Then AddCustomSubButton(sub_Edit, "Freeze", "Scheme_IsFrozen", faceId_Freeze).BeginGroup = True
        If .Can_Undelete Then AddCustomSubButton sub_Edit, "Undelete", "Scheme_Undelete", faceId_Undelete
        If .Can_RenumberWBS Then AddCustomSubButton sub_Edit, "Renumber WBS", "Scheme_RenumberWbs", faceId_WbsRenumber
    End With

End Sub

Private Sub SetDefaults_Create()
    Dim sc As String: sc = Form_frm_Schemes.cmb_Main
    Me.schemeCode = sc
    Me.indent = Nz(DLast("indent", "tbl_" & suffix, "schemeCode='" & sc & "'"), 0)
    Me.sort = Nz(DMax("sort", "tbl_" & suffix, "schemeCode='" & sc & "'"), 0) + 1
    Me.sysStartPerson = g_Model.User_Current
    Me.object = prefix & Me.dir
    Me.sysGuid = CreateGuid
    Me.Title = TreeIndent_Set(Me.Title, Me.indent)
End Sub

Private Sub SetDefaults_Update()
    Me.sysStartTime = Now
    Me.sysStartPerson = g_Model.User_Current
    Me.sysGuid = CreateGuid
End Sub

' ------ '
' PUBLIC '
' ------ '

Public Property Get Prefix_() As String
    Prefix_ = prefix
End Property

