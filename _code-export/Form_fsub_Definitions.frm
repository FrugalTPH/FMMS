VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_fsub_Definitions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const prefix As String = "D"
Private Const suffix As String = "Definitions"
Private s_recordsForDeletion As String
Private ts_recordsForDeletion As Date


Private Sub code_BeforeUpdate(Cancel As Integer)
    If Me.code.OldValue <> Me.code Then
        MsgBox "Definition codes cannot be changed once created", vbInformation, "Code change denied"
        Me.Undo
        Cancel = True
    End If
End Sub

Private Sub code_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = acRightButton Then
        If Nz(Me!code, vbNullString) <> vbNullString And Nz(Me!sysEndPerson, 0) > 0 Then
            Refresh_ContextMenu
            CommandBars("menu_" & suffix).ShowPopup
            DoCmd.CancelEvent
        End If
    End If
End Sub

Private Sub Form_AfterDelConfirm(status As Integer)
    If status <> acDeleteOK Then
        Dim r As Variant
        For Each r In Split(s_recordsForDeletion, ",")
            g_Model.Definition_Update Me!field, CLng(r), vbNullString, ts_recordsForDeletion
        Next r
    End If
    s_recordsForDeletion = vbNullString
    ts_recordsForDeletion = vbMaxDate
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
        Case Nz(Me!code, vbNullString) = vbNullString:
            strTitle = strTitle & "Invalid Code"
            strMsg = "Code cannot be null."
        Case Not g_Model.Definition_Archive("tbl_" & suffix, Me!code):
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
    Form_frm_Definitions.Refresh_Captions
End Sub

Private Sub Form_Delete(Cancel As Integer)
    If ts_recordsForDeletion = vbMaxDate Then ts_recordsForDeletion = Now
    s_recordsForDeletion = s_recordsForDeletion & Me!code & ","
    If g_Model.Definition_Archive("tbl_" & suffix, Me!code) Then Exit Sub
abort:
    MsgBox "The record couldn't be archived, so the deletion cannot proceed.", vbExclamation, "Delete Denied - Archiving Failed"
    Cancel = True
End Sub

Private Sub Form_Load()
    ts_recordsForDeletion = vbMaxDate
End Sub

Private Sub Form_Open(Cancel As Integer)
    Query_Refresh "quni_" & suffix, std_Sql.Definitions_quni
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

Private Sub sysEndPerson_GotFocus()
    sysEndPerson.Requery
End Sub

Private Sub sysStartPerson_GotFocus()
    sysStartPerson.Requery
End Sub

Private Sub Title_BeforeUpdate(Cancel As Integer)
    Cancel = IsNull(Me.Title)
End Sub


' ------- '
' PRIVATE '
' ------- '

Private Sub Refresh_ContextMenu()
    
    On Error Resume Next
        CommandBars("menu_" & suffix).Delete
    On Error GoTo 0
    
    Dim cbar As CommandBar: Set cbar = CommandBars.Add(name:="menu_" & suffix, Position:=msoBarPopup)
    AddCustomButton cbar, "Undelete", "Definition_Undelete", faceId_Undelete
    
End Sub

Private Sub SetDefaults_Create()
    Dim field As String: field = Form_frm_Definitions.cmb_Main
    Me.field = field
    Me.sysStartPerson = g_Model.User_Current
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
