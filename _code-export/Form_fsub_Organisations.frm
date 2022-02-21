VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_fsub_Organisations"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const prefix As String = "O"
Private Const suffix As String = "Organisations"
Private s_recordsForDeletion As String
Private ts_recordsForDeletion As Date


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
        Case Nz(Me.identity, vbNullString) <> vbNullString And Not g_Model.User_IsManager:
            strTitle = strTitle & "Insufficient Permissions"
            strMsg = "Only managers can update the details of organisations with an originator code (stored in the identity column)."
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
    Form_frm_Organisations.Refresh_Captions
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
    Query_Refresh "quni_" & suffix, std_Sql.Organisations_quni
    Me.RecordSource = "qry_" & suffix
End Sub

Private Sub identity_BeforeUpdate(Cancel As Integer)
    If Not g_Model.User_IsManager Then
        Me.Undo
        Cancel = True
    End If
End Sub

Private Sub identity_GotFocus()
    identity.Requery
End Sub

Private Sub link_AfterUpdate()
    If IsNull(link) Then Exit Sub
    If link = vbNullString Then Exit Sub
    Dim URL As String: URL = mod_StringUtils.HyperlinkMidPart(link)
    link = "Open#" & URL & "##" & URL
End Sub

Private Sub Link_Click()
    LinkClick Me
End Sub

Private Sub sysEndPerson_GotFocus()
    sysEndPerson.Requery
End Sub

Private Sub sysStartPerson_GotFocus()
    sysStartPerson.Requery
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
    
        If .Can_ReadFs Then
            AddCustomButton cbar, "Open (WIP)", "Organisation_OpenWIP", faceId_WorkInProgress
            AddCustomButton cbar, "Latest Snapshot", "Organisation_OpenLatestSnapshot", faceId_SnapshotLatest
            AddCustomButton cbar, "All Snapshots", "Organisation_OpenSnapshots", faceId_SnapshotAll
        End If
        AddCustomButton cbar, "Comments", "Organisation_Comments", faceId_Comments
            
        Dim sub_Edit As CommandBarControl: Set sub_Edit = cbar.Controls.Add(Type:=msoControlPopup)
        sub_Edit.Caption = "Edit"
        sub_Edit.Enabled = .Can_EditModel
        If .Can_Attach Then AddCustomSubButton(sub_Edit, "Attach => " & .scheme, "Organisation_Attach", faceId_Attach).BeginGroup = True
        If .Can_Detach Then AddCustomSubButton(sub_Edit, "Detach <= " & .scheme, "Organisation_Detach", faceId_Detach).BeginGroup = True
        If .Can_TakeSnapshot Then AddCustomSubButton sub_Edit, "Take Snapshot", "Organisation_CreateSnapshot", faceId_SnapshotSave
        If .Can_Unfreeze Then AddCustomSubButton(sub_Edit, "Unfreeze", "Organisation_IsFrozen", faceId_Unfreeze).BeginGroup = True
        If .Can_Freeze Then AddCustomSubButton(sub_Edit, "Freeze", "Organisation_IsFrozen", faceId_Freeze).BeginGroup = True
        If .Can_Undelete Then AddCustomSubButton sub_Edit, "Undelete", "Organisation_Undelete", faceId_Undelete

    End With
    
End Sub

Private Sub SetDefaults_Create()
    Me.sysStartPerson = g_Model.User_Current
    Me.object = prefix & Me.dir
    Me.sysGuid = CreateGuid
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
