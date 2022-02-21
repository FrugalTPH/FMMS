VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frm_Definitions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const prefix As String = "D"
Private Const suffix As String = "Definitions"
Public w As Long
Public h As Long



Private Sub btn_PublishDb_Click()
    g_Model.PublishDb
End Sub

Private Sub btn_PublishSheet_Click()
    g_Model.PublishSheet
End Sub

Private Sub chk_Deleted_AfterUpdate()
    Refresh_RecordSource
End Sub

Private Sub chk_readOnly_DblClick(Cancel As Integer)
    If chk_readOnly Then
        If MsgBox("Click OK to make the model editable?" & vbCrLf & vbCrLf & "Note: The application will close and need to be restarted for this to take effect.", vbOKCancel, "Unlock model") <> vbOK Then Exit Sub
        g_Model.Db_ValueSet "is-editable", "true", True
    Else
        If MsgBox("Click OK to make the model read-only?" & vbCrLf & vbCrLf & "Note: The application will close and need to be restarted for this to take effect.", vbOKCancel, "Lock model") <> vbOK Then Exit Sub
        g_Model.Db_ValueSet "is-editable", "false"
    End If
    Form_frm_Schemes.ForceClose
End Sub

Private Sub cmb_Main_AfterUpdate()
    sub_Browser.Requery
End Sub

Private Sub cmb_Main_KeyDown(KeyCode As Integer, Shift As Integer)
    KeyCode = 0
End Sub

Private Sub Form_Activate()
    cmb_Main.Requery
    sub_Browser.SetFocus
End Sub

Private Sub Form_Close()
    Form_frm_Schemes.btn_Definitions.BackStyle = 0
End Sub

Private Sub Form_Load()
    
    Set_FormIcon Me, LCase$(suffix)
    Set_FormPermissions Me.Form, Form_fsub_Schemes
    Set_FormSize Me

    FormHeader.BackColor = Solid.Black
    Detail.BackColor = Solid.Black
    FormFooter.BackColor = Solid.Black
    
    cmb_Main.RowSource = std_Sql.cmb_Main_D
    cmb_Main = cmb_Main.Column(0, 0)
    
    chk_readOnly = g_Model.Db_isReadOnly
    
    Refresh_Captions
    Refresh_RecordSource
    Form_frm_Schemes.btn_Definitions.BackStyle = 1
    
End Sub

Private Sub Form_Resize()
    Limit_FormSize Me
End Sub

Private Property Get Form_Title() As String
    Form_Title = "Definitions   ~   No selection"
    Dim frm As Form: Set frm = sub_Browser.Form
    If frm.Recordset.RecordCount > 0 Then
        Dim code As String: code = Nz(frm!code, vbNullString)
        Dim Title As String: Title = Nz(frm!Title, vbNullString)
        If code = vbNullString Or Title = vbNullString Then Exit Property
        Form_Title = "Definitions   ~   " & code & "  /  " & mod_StringUtils.StripNonAsciiChars(Title)
    End If
End Property


' ------- '
' PRIVATE '
' ------- '

Private Sub Refresh_RecordSource()
    Dim strSql As String
    If chk_Deleted Then
        strSql = std_Sql.Definitions_old
    Else
        strSql = std_Sql.Definitions_current
    End If
    Query_Refresh "qry_" & suffix, strSql
    sub_Browser.Form.RecordSource = "qry_" & suffix
    Refresh_Captions
End Sub

' ------ '
' PUBLIC '
' ------ '

Public Property Get Prefix_() As String
    Prefix_ = prefix
End Property

Public Sub Refresh_Captions()
    Caption = Form_Title
End Sub
