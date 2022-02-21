VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_fsub_Outputs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const prefix As String = "U"
Private Const suffix As String = "Outputs"
Private s_recordsForDeletion As String
Private ts_recordsForDeletion As Date


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
    Form_frm_Outputs.Refresh_Captions
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
    Query_Refresh "quni_" & suffix, std_Sql.Outputs_quni
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
    Auto_PersonByScheme Me, Nz(Me!scheme, 0)
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
    Auto_ClassifyByScheme Me
End Sub

Private Sub revCode_GotFocus()
    revCode.Requery
End Sub

Private Sub revNo_AfterUpdate()
    If Not IsNull(revNo) And Len(revNo) > 0 Then revNo = ValidDocRevNo(revNo)
End Sub

Private Sub revNo_DblClick(Cancel As Integer)
    If g_Model.Db_isReadOnly Then Exit Sub
    If g_Model.Fs_isConnected Then
        Revise_Current
    Else
        MsgBox "Filesystem access is required to be able to revise Outputs.", vbInformation, "Revise Output - Denied"
    End If
End Sub

Private Sub revNo_Enter()
    If g_Model.Db_isReadOnly Then Exit Sub
    With Me
        If Nz(!revNo, vbNullString) <> vbNullString Then Exit Sub
        If Nz(!revCode, vbNullString) = vbNullString Then Exit Sub
        If Nz(!refNo, vbNullString) = vbNullString Then Exit Sub
        Dim revSuffix As String: revSuffix = GetRevNo_Suffix(!refNo, !revCode)
        If revSuffix <> vbNullString Then !revNo = !revCode & revSuffix
    End With
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

Private Sub scheme_GotFocus()
    scheme.Requery
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

Private Sub typeCode_Enter()
    Auto_ClassifyByScheme Me
End Sub

Private Sub typeCode_GotFocus()
    TypeCode.Requery
End Sub

Private Sub wsCode_DblClick(Cancel As Integer)
    Manual_Classify Me
End Sub

Private Sub wsCode_Enter()
    Auto_ClassifyByScheme Me
End Sub

Private Sub wsCode_GotFocus()
    wsCode.Requery
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
    With dbLocal.OpenRecordset("SELECT DISTINCT refNo FROM quni_Outputs WHERE refNo ALike '" & refPrefix & "%';")
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

Private Function GetRevNo_Suffix(strRefNo As String, revPrefix As String) As String
    Dim strPd As String: strPd = g_Model.Db_ValueGet("rev-padding")
    If strPd = vbNullString Then Exit Function
    Dim pad As String: pad = String(CLng(strPd), "0")
    Dim maxVal As Long: maxVal = 0
    Dim curVal As Long: curVal = 0
    With dbLocal.OpenRecordset("SELECT DISTINCT revNo FROM quni_Outputs WHERE refNo = '" & strRefNo & "' AND revNo ALike '" & revPrefix & "%';")
        If Not (.EOF And .BOF) Then
            .MoveFirst
            Do Until .EOF = True
                curVal = val(StripToTheseCharacters(!revNo, "0123456789"))
                If curVal > maxVal Then maxVal = curVal
                .MoveNext
            Loop
        End If
        .Close
    End With
    GetRevNo_Suffix = Format$(maxVal + 1, pad)
End Function

Private Sub Refresh_ContextMenu()
    
    On Error Resume Next
        CommandBars("menu_" & suffix).Delete
    On Error GoTo 0
    
    Dim cbar As CommandBar: Set cbar = CommandBars.Add(name:="menu_" & suffix, Position:=msoBarPopup)
        
    With New cls_Permissor
        
        .Init Me.Form
    
        If .Can_ReadFs Then
            AddCustomButton cbar, "Open (SHARED)", "Output_OpenWIP", faceId_WorkInProgress
            AddCustomButton cbar, "Latest Snapshot", "Output_OpenLatestSnapshot", faceId_SnapshotLatest
            AddCustomButton cbar, "All Snapshots", "Output_OpenSnapshots", faceId_SnapshotAll
        End If
        
        AddCustomButton cbar, "Comments", "Output_Comments", faceId_Comments
        
        Dim sub_Edit As CommandBarControl: Set sub_Edit = cbar.Controls.Add(Type:=msoControlPopup)
        sub_Edit.Caption = "Edit"
        sub_Edit.Enabled = .Can_EditModel
        If .Can_Attach Then AddCustomSubButton(sub_Edit, "Attach => " & .scheme, "Output_Attach", faceId_Attach).BeginGroup = True
        If .Can_Detach Then AddCustomSubButton(sub_Edit, "Detach <= " & .scheme, "Output_Detach", faceId_Detach).BeginGroup = True
        If .Can_TakeSnapshot Then AddCustomSubButton sub_Edit, "Take Snapshot", "Output_CreateSnapshot", faceId_SnapshotSave
        If .Can_Duplicate Then AddCustomSubButton sub_Edit, "Duplicate", "Output_Duplicate", faceId_Duplicate
        If .Can_Revise Then AddCustomSubButton sub_Edit, "Revise", "Output_Revise", faceId_Revise
        If .Can_Unfreeze Then AddCustomSubButton(sub_Edit, "Unfreeze", "Output_IsFrozen", faceId_Unfreeze).BeginGroup = True
        If .Can_Freeze Then AddCustomSubButton(sub_Edit, "Freeze", "Output_IsFrozen", faceId_Freeze).BeginGroup = True
        If .Can_Undelete Then AddCustomSubButton sub_Edit, "Undelete", "Output_Undelete", faceId_Undelete
        
    End With
    
End Sub

Private Sub SetDefaults_Create()
    Me.scheme = Form_frm_Schemes.CurrentScheme
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

Public Sub Revise_Current()
    If Nz(Me!revCode, vbNullString) = vbNullString Then Exit Sub
    If MsgBox("OK to take a snapshot and increment the revision now?", vbOKCancel + vbQuestion, "Increment Revision") <> vbOK Then Exit Sub
    With Me
        If Nz(!revCode, vbNullString) = vbNullString Then Exit Sub
        If Nz(!refNo, vbNullString) = vbNullString Then Exit Sub
        g_Model.Snapshot_Create prefix, !dir, !refNo & "_" & !revNo
        Dim revSuffix As String: revSuffix = GetRevNo_Suffix(!refNo, !revCode)
        If revSuffix <> vbNullString Then !revNo = !revCode & revSuffix
        .Dirty = False
    End With
End Sub
