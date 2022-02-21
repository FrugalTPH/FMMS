VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_fsub_Inputs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const prefix As String = "N"
Private Const suffix As String = "Inputs"
Private s_recordsForDeletion As String
Private ts_recordsForDeletion As Date


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

Private Sub effectiveOrg_GotFocus()
    effectiveOrg.Requery
End Sub

Private Sub effectiveOrg_NotInList(newData As String, Response As Integer)
    Response = Organisation_CreateNotInList(effectiveOrg, newData)
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
    Cancel = MsgBox("OK to bypass the Filestage and register a new empty Input?", vbOKCancel + vbQuestion, "Bypass Filestage?") <> vbOK
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
    Form_frm_Inputs.Refresh_Captions
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
    Query_Refresh "quni_" & suffix, std_Sql.Inputs_quni
    Me.RecordSource = "qry_" & suffix
End Sub

Private Sub isTemplate_BeforeUpdate(Cancel As Integer)
    If Not g_Model.User_IsManager Then
        MsgBox "Only managers can define document templates", vbInformation, "Update isTemplate - Denied!"
        isTemplate.Undo
        Cancel = True
    End If
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

Private Sub refNo_AfterUpdate()
    If Not IsNull(refNo) And Len(refNo) > 0 Then refNo = ValidDocRefNo(refNo)
End Sub

Private Sub revNo_AfterUpdate()
    If Not IsNull(revNo) And Len(revNo) > 0 Then revNo = ValidDocRevNo(revNo)
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

Private Sub sysEndPerson_GotFocus()
    sysEndPerson.Requery
End Sub

Private Sub sysStartPerson_GotFocus()
    sysStartPerson.Requery
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
            AddCustomButton cbar, "Open (SHARED)", "Input_OpenWIP", faceId_WorkInProgress
            AddCustomButton cbar, "Latest Snapshot", "Input_OpenLatestSnapshot", faceId_SnapshotLatest
            AddCustomButton cbar, "All Snapshots", "Input_OpenSnapshots", faceId_SnapshotAll
        End If
        
        AddCustomButton cbar, "Comments", "Input_Comments", faceId_Comments
        
        Dim sub_Edit As CommandBarControl: Set sub_Edit = cbar.Controls.Add(Type:=msoControlPopup)
        sub_Edit.Caption = "Edit"
        sub_Edit.Enabled = .Can_EditModel
        If .Can_Attach Then AddCustomSubButton(sub_Edit, "Attach => " & .scheme, "Input_Attach", faceId_Attach).BeginGroup = True
        If .Can_Detach Then AddCustomSubButton(sub_Edit, "Detach <= " & .scheme, "Input_Detach", faceId_Detach).BeginGroup = True
        If .Can_TakeSnapshot Then AddCustomSubButton sub_Edit, "Take Snapshot", "Input_CreateSnapshot", faceId_SnapshotSave
        If .Can_Unfreeze Then AddCustomSubButton(sub_Edit, "Unfreeze", "Input_IsFrozen", faceId_Unfreeze).BeginGroup = True
        If .Can_Freeze Then AddCustomSubButton(sub_Edit, "Freeze", "Input_IsFrozen", faceId_Freeze).BeginGroup = True
        If .Can_Undelete Then AddCustomSubButton sub_Edit, "Undelete", "Input_Undelete", faceId_Undelete

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
