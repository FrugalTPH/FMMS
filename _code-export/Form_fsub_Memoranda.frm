VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_fsub_Memoranda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const prefix As String = "M"
Private Const suffix As String = "Memoranda"
Private s_recordsForDeletion As String
Private ts_recordsForDeletion As Date


Private Sub accountable_Enter()
    Auto_PersonByScheme Me, 0
End Sub

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
    Auto_ClassifyByScheme Me
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
    Form_frm_Memoranda.Refresh_Captions
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
    Query_Refresh "quni_" & suffix, std_Sql.Memoranda_quni
    Me.RecordSource = "qry_" & suffix
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
    Auto_ClassifyByScheme Me
End Sub

Private Sub locCode_GotFocus()
    locCode.Requery
End Sub

Private Sub origCode_DblClick(Cancel As Integer)
    Manual_Classify Me
End Sub

Private Sub origCode_Enter()
    Auto_ClassifyByScheme Me
End Sub

Private Sub origCode_GotFocus()
    origCode.Requery
End Sub

Private Sub refNo_AfterUpdate()
    If Not IsNull(refNo) And Len(refNo) > 0 Then refNo = ValidDocRefNo(refNo)
End Sub

Private Sub refNo_Enter()
    If g_Model.Db_isReadOnly Then Exit Sub
    If Nz(Me!refNo, vbNullString) <> vbNullString Then Exit Sub
    Dim refPrefix As String: refPrefix = GetRefNo_Prefix
    If refPrefix = vbNullString Then Exit Sub
    Dim refSuffix As String: refSuffix = GetRefNo_Suffix(refPrefix)
    If refSuffix <> vbNullString Then Me!refNo = refPrefix & refSuffix
End Sub

Private Sub responsible_Enter()
    Auto_PersonByScheme Me, 0
End Sub

Private Sub responsible_GotFocus()
    responsible.Requery
End Sub

Private Sub responsible_NotInList(newData As String, Response As Integer)
    Response = Person_CreateNotInList(responsible, newData)
End Sub

Private Sub roleCode_DblClick(Cancel As Integer)
    Manual_Classify Me
End Sub

Private Sub roleCode_Enter()
    Auto_ClassifyByScheme Me
End Sub

Private Sub roleCode_GotFocus()
    roleCode.Requery
End Sub

Private Sub secCode_DblClick(Cancel As Integer)
    Manual_Classify Me
End Sub

Private Sub secCode_Enter()
    Auto_ClassifyByScheme Me
End Sub

Private Sub secCode_GotFocus()
    secCode.Requery
End Sub

Private Sub statusCode_DblClick(Cancel As Integer)
    Manual_Classify Me
End Sub

Private Sub statusCode_Enter()
    Auto_ClassifyByScheme Me
End Sub

Private Sub statusCode_GotFocus()
    statusCode.Requery
End Sub

Private Sub sysCode_DblClick(Cancel As Integer)
    Manual_Classify Me
End Sub

Private Sub sysCode_Enter()
    Auto_ClassifyByScheme Me
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

Private Sub Title_BeforeUpdate(Cancel As Integer)
    Cancel = IsNull(Me.Title)
End Sub

Private Sub typeCode_DblClick(Cancel As Integer)
    Manual_Template Me
End Sub

Private Sub typeCode_GotFocus()
    TypeCode.Requery
End Sub

' ------- '
' PRIVATE '
' ------- '

Private Function GetRefNo_Prefix() As String
    With Me
        Dim projectCode As String: projectCode = TryGetClassification("projectCode", prefix)
        Select Case True
            Case Nz(projectCode, vbNullString) = vbNullString: Exit Function
            Case Nz(!origCode, vbNullString) = vbNullString: Exit Function
            Case Nz(!sysCode, vbNullString) = vbNullString: Exit Function
            Case Nz(!locCode, vbNullString) = vbNullString: Exit Function
            Case Nz(!TypeCode, vbNullString) = vbNullString: Exit Function
            Case Nz(!roleCode, vbNullString) = vbNullString: Exit Function
            Case Else: GetRefNo_Prefix = Join(Array(projectCode, !origCode, !sysCode, !locCode, !TypeCode, !roleCode), vbDash) & vbDash
        End Select
    End With
End Function

Private Function GetRefNo_Suffix(refPrefix As String) As String
    Dim strPd As String: strPd = g_Model.Db_ValueGet("ref-padding")
    If strPd = vbNullString Then Exit Function
    Dim pad As String: pad = String(CLng(strPd), "0")
    Dim maxVal As Long: maxVal = 0
    Dim curVal As Long: curVal = 0
    With dbLocal.OpenRecordset("SELECT DISTINCT refNo FROM quni_Memoranda WHERE refNo ALike '" & refPrefix & "%';")
        If Not (.EOF And .BOF) Then
            .MoveFirst
            Do Until .EOF = True
                curVal = RightNumberPart(!refNo, "-")
                If curVal > maxVal Then maxVal = curVal
                .MoveNext
            Loop
        End If
        .Close
    End With
    GetRefNo_Suffix = Format$(maxVal + 1, pad)
End Function

Private Sub Refresh_ContextMenu()
    
    On Error Resume Next
        CommandBars("menu_" & suffix).Delete
    On Error GoTo 0
    
    Dim cbar As CommandBar: Set cbar = CommandBars.Add(name:="menu_" & suffix, Position:=msoBarPopup)
        
    With New cls_Permissor
        
        .Init Me.Form
    
        If .Can_ReadFs Then
            AddCustomButton cbar, "Open (SHARED)", "Memorandum_OpenWIP", faceId_WorkInProgress
            AddCustomButton cbar, "Latest Snapshot", "Memorandum_OpenLatestSnapshot", faceId_SnapshotLatest
            AddCustomButton cbar, "All Snapshots", "Memorandum_OpenSnapshots", faceId_SnapshotAll

        End If
        
        AddCustomButton cbar, "Comments", "Memorandum_Comments", faceId_Comments
        AddCustomButton cbar, "Uncertainty", "Memorandum_Uncertainty", faceId_Uncertainty
        
        Dim sub_Edit As CommandBarControl: Set sub_Edit = cbar.Controls.Add(Type:=msoControlPopup)
        sub_Edit.Caption = "Edit"
        sub_Edit.Enabled = .Can_EditModel
        If .Can_Attach Then AddCustomSubButton(sub_Edit, "Attach => " & .scheme, "Memorandum_Attach", faceId_Attach).BeginGroup = True
        If .Can_Detach Then AddCustomSubButton(sub_Edit, "Detach <= " & .scheme, "Memorandum_Detach", faceId_Detach).BeginGroup = True
        If .Can_TakeSnapshot Then AddCustomSubButton sub_Edit, "Take Snapshot", "Memorandum_CreateSnapshot", faceId_SnapshotSave
        If .Can_Duplicate Then AddCustomSubButton sub_Edit, "Duplicate", "Memorandum_Duplicate", faceId_Duplicate
        If .Can_Unfreeze Then AddCustomSubButton(sub_Edit, "Unfreeze", "Memorandum_IsFrozen", faceId_Unfreeze).BeginGroup = True
        If .Can_Freeze Then AddCustomSubButton(sub_Edit, "Freeze", "Memorandum_IsFrozen", faceId_Freeze).BeginGroup = True
        If .Can_Undelete Then AddCustomSubButton sub_Edit, "Undelete", "Memorandum_Undelete", faceId_Undelete
        
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
