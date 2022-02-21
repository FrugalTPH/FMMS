VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_fdlg_Commenter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private s_obj As String
Private Const suffix As String = "Comments"
Public w As Long
Public h As Long


Private Sub btn_Submit_Click()
    g_Model.Comment_Create s_obj, txt_Comment
    txt_Comment = vbNullString
    txt_Comment.SetFocus
On Error Resume Next
    With sub_Browser.Form
        .Requery
        .comment.SelStart = 0
    End With
End Sub

Private Sub Form_Activate()
    txt_Comment.SetFocus
End Sub

Private Sub Form_Load()

    Set_FormIcon Me, LCase$(suffix)
    Set_FormSize Me
    
    FormHeader.BackColor = Solid.White
    Detail.BackColor = Solid.White
    FormFooter.BackColor = Solid.White
    
    Refresh_RecordSource Replace(std_Sql.Comments_1, "{0}", s_obj)
    
    Caption = "Comments   ~   " & s_obj
    
On Error Resume Next
    sub_Browser.Form!comment.SelStart = 0
End Sub

Private Sub Form_Open(Cancel As Integer)

    s_obj = Nz(Me.OpenArgs, vbNullString)
    If s_obj = vbNullString Then DoCmd.Close acForm, Me.name, acSaveNo

    txt_Comment.Locked = g_Model.Db_isReadOnly
    btn_Submit.Enabled = Not txt_Comment.Locked
    
End Sub

Private Sub Form_Resize()
    Limit_FormSize Me
End Sub

' ------- '
' PRIVATE '
' ------- '

' ------ '
' PUBLIC '
' ------ '

Public Sub Refresh_RecordSource(strSql As String)
    Query_Refresh "qry__Comments", strSql
    sub_Browser.Form.RecordSource = "qry__Comments"
End Sub
