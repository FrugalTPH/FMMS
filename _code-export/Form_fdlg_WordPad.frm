VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_fdlg_WordPad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private state As Integer                ' 0 = Undefined, -1 = read-only, 1 = read-write
Private old_wp As String
Private s_FormName As String
Private s_object As String
Private s_prefix As String
Public w As Long
Public h As Long


Private Sub btn_Cancel_Click()
    DoCmd.Close acForm, Me.name, acSaveYes
End Sub

Private Sub btn_LoadTemplate_Click()
    If MsgBox("Ok to reset current WordPad content to the model default?" & vbCrLf & "Note: This won't be committed until 'Save Changes' is executed.", vbOKCancel, "Load Template ~ " & s_prefix) <> vbOK Then Exit Sub
    wp = g_Model.Db_ValueGet("wp-" & s_prefix)
End Sub

Private Sub btn_SaveChanges_Click()
    If Not IsChanged Then
        MsgBox "There have been no changes since last revision, so no save is required.", vbInformation, "No changes detected"
        Exit Sub
    End If
    If MsgBox("Ok to save these changes as last version?", vbOKCancel, "Save Changes") <> vbOK Then Exit Sub
    If g_Model.WordPad_ArchiveAndUpdate(s_object, "wp=""" & wp & """") Then
        cmb_Main.Requery
        cmb_Main = cmb_Main.Column(0, 0)
        Refresh_Controls
    End If
End Sub

Private Sub btn_SaveTemplate_Click()
    Dim strKey As String: strKey = Left(s_object, 1)
    If MsgBox("Ok to save this WordPad text as the default?", vbOKCancel, "Overwrite Template ~ " & strKey) <> vbOK Then Exit Sub
    g_Model.Db_ValueSet "wp-" & strKey, wp
End Sub

Private Sub btn_Undo_Click()
    If Not IsChanged Then
        MsgBox "There have been no changes since last revision, so no need to undo.", vbInformation, "No changes detected"
        Exit Sub
    End If
    If MsgBox("Ok to revert to the last saved version?", vbOKCancel, "Undo Changes") <> vbOK Then Exit Sub
    Refresh_Controls
    cmb_Main.Requery
    cmb_Main = cmb_Main.Column(0, 0)
End Sub

Private Sub cmb_Main_AfterUpdate()
    Refresh_Controls
End Sub

Private Sub Form_Load()
    Set_FormIcon Me, "wordpad"
    Set_FormSize Me
    cmb_Main.RowSource = Replace(std_Sql.cmb_Main_W, "{0}", s_object)
    cmb_Main = cmb_Main.Column(0, 0)
    Refresh_Controls
End Sub

Private Sub Form_Open(Cancel As Integer)

    s_FormName = Nz(Me.OpenArgs, vbNullString)
    If s_FormName = vbNullString Then DoCmd.Close acForm, Me.name, acSaveNo
    
    Dim frm As Form: Set frm = GetForm(s_FormName)
    If frm Is Nothing Then DoCmd.Close acForm, Me.name, acSaveNo
    
    s_object = frm.sub_Browser!object
    s_prefix = Left(s_object, 1)

    FormHeader.BackColor = frm.FormHeader.BackColor
    Detail.BackColor = frm.Detail.BackColor
    FormFooter.BackColor = frm.FormFooter.BackColor
    
    If ECount("quni_WordPad", "object='" & s_object & "'") <= 0 Then g_Model.WordPad_Create s_object
    
End Sub

Private Sub Form_Resize()
    Limit_FormSize Me
End Sub

' ------- '
' PRIVATE '
' ------- '

Private Function GetData() As String
    Dim rst As Recordset: Set rst = dbLocal.OpenRecordset("SELECT * FROM quni_WordPad WHERE object='" & s_object & "'", dbOpenDynaset)
    rst.FindFirst "Format$(sysStartTime, 'dd\/mm\/yyyy hh\:nn\:ss') = '" & SelectedDate & "'"
    GetData = Nz(rst!wp, vbNullString)
    rst.Close
    Set rst = Nothing
End Function

Private Function IsChanged() As Boolean
    IsChanged = old_wp <> wp
End Function

Private Sub Refresh_Controls()

    TrySetState
        
    Caption = "WordPad (current)   ~   " & s_object
    
    ' read-write
    If state > 0 Then
        wp.Locked = False
        btn_SaveTemplate.Visible = True
        btn_LoadTemplate.Visible = True
        btn_SaveChanges.Visible = True
        btn_Undo.Visible = True
    End If
    
    ' read-only
    If state < 0 Or cmb_Main.ListIndex > 0 Then
        wp.Locked = True
        btn_SaveTemplate.Visible = False
        btn_LoadTemplate.Visible = False
        btn_SaveChanges.Visible = False
        btn_Undo.Visible = False
    End If
        
    wp = GetData
    old_wp = wp
    
    If cmb_Main.ListIndex > 0 Then Caption = Replace(Caption, "current", "old")
    
End Sub

Private Sub TrySetState()
    If state = 0 Then
        Dim bln_IsReadOnly As Boolean: bln_IsReadOnly = g_Model.Db_isReadOnly
        Dim bln_IsManager As Boolean: bln_IsManager = g_Model.User_IsManager
        Dim frm As Form: Set frm = GetForm(s_FormName)
        Dim bln_isFrozen As Boolean: bln_isFrozen = frm.sub_Browser!isFrozen
        Select Case True
            Case bln_IsReadOnly: state = -1
            Case bln_isFrozen: state = -1
            Case Not bln_IsManager: state = -1
            Case Else: state = 1
        End Select
    End If
End Sub

Private Property Get SelectedDate() As String
    SelectedDate = Format$(cmb_Main, "dd\/mm\/yyyy hh\:nn\:ss")
End Property

' ------ '
' PUBLIC '
' ------ '
