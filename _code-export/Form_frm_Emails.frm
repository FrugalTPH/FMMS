VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frm_Emails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const prefix As String = "E"
Private Const suffix As String = "Emails"
Public w As Long
Public h As Long


Private Sub chk_Children_AfterUpdate()
    Refresh_RecordSource
End Sub

Private Sub chk_Deleted_AfterUpdate()
    Refresh_RecordSource
End Sub

Private Sub chk_Parents_AfterUpdate()
    Refresh_RecordSource
End Sub

Private Sub chk_Selected_AfterUpdate()
    Refresh_RecordSource
End Sub

Private Sub cmb_Main_AfterUpdate()
    Refresh_RecordSource
    Refresh_Controls
End Sub

Private Sub cmb_Main_KeyDown(KeyCode As Integer, Shift As Integer)
    KeyCode = 0
End Sub

Private Sub Form_Activate()
    cmb_Main.Requery
    sub_Browser.SetFocus
End Sub

Private Sub Form_Close()
    Form_frm_Schemes.btn_Emails.BackStyle = 0
End Sub

Private Sub Form_Load()
    
    Set_FormIcon Me, LCase$(suffix)
    Set_FormPermissions Me.Form, Form_fsub_Schemes
    Set_FormSize Me
    
    FormHeader.BackColor = Pastel_0.Grape
    Detail.BackColor = Pastel_0.Grape
    FormFooter.BackColor = Pastel_0.Grape
   
    cmb_Main.RowSource = std_Sql.cmb_Main_E
    cmb_Main = cmb_Main.Column(1, 0)
    
    chk_Parents = False
    chk_Selected = True
    chk_Children = False
    chk_Deleted = False
    
    Refresh_Controls
    Refresh_RecordSource
    Form_frm_Schemes.btn_Emails.BackStyle = 1
    
End Sub

Private Sub Form_Resize()
    Limit_FormSize Me
End Sub

Private Property Get Form_Subject() As String
    Form_Subject = "Emails   ~   No selection"
    Dim frm As Form: Set frm = sub_Browser.Form
    If frm.Recordset.RecordCount > 0 Then
        Dim dir As Long: dir = Nz(frm!dir, 0)
        Dim subject As String: subject = Nz(frm!subject, vbNullString)
        If dir = 0 Or subject = vbNullString Then Exit Property
        Form_Subject = "Emails   ~   " & prefix & dir & "  /  " & mod_StringUtils.StripNonAsciiChars(subject)
    End If
End Property

Private Sub lv_DropZone_OLEDragDrop(Data As Object, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim i As Long
    Const vbCFFiles = 15
    If Data.GetFormat(vbCFFiles) Then
        For i = 1 To Data.files.count
            Debug.Print Data.files(i)
        Next i
    Else
        Debug.Print "No file(s) dropped."
    End If
End Sub

Private Sub lbl_Mailstage_Click()
    DoCmd.OpenForm "frm_Mailstage"
End Sub

' ------- '
' PRIVATE '
' ------- '

' ------ '
' PUBLIC '
' ------ '

Public Property Get Prefix_() As String
    Prefix_ = prefix
End Property

Public Sub Refresh_Captions()
    Caption = Form_Subject
    If cmb_Main > 1 Then
        FormHeader.BackColor = Pastel_0.Gray
        Detail.BackColor = Pastel_0.Gray
        FormFooter.BackColor = Pastel_0.Gray
    Else
        FormHeader.BackColor = Pastel_0.Grape
        Detail.BackColor = Pastel_0.Grape
        FormFooter.BackColor = Pastel_0.Grape
    End If
End Sub

Public Sub Refresh_Controls()
    chk_Deleted.Visible = cmb_Main = 1
    chk_Children.Visible = cmb_Main = 2
    chk_Parents.Visible = cmb_Main = 2
    chk_Selected.Visible = cmb_Main = 2
End Sub

Public Sub Refresh_RecordSource()
    Dim strSql As String
    If cmb_Main = 1 Then
        If chk_Deleted Then
            strSql = std_Sql.Emails_old
        Else
            strSql = std_Sql.Emails_current
        End If
    Else
        strSql = std_Sql.Emails_byCriteria
        Dim criteria As String: criteria = vbNullString
        With Form_frm_Schemes
            If cmb_Main = 2 Then
                If chk_Children Then criteria = criteria & vbCrLf & .Get_Criteria(Sql_Atta_Child)
                If chk_Parents Then criteria = criteria & vbCrLf & .Get_Criteria(Sql_Atta_Parent)
                If chk_Selected Then criteria = criteria & vbCrLf & .Get_Criteria(Sql_Atta_Scheme)
            End If
        End With
        strSql = Replace(strSql, "{0}", criteria)
    End If
    Query_Refresh "qry_" & suffix, strSql
    sub_Browser.Form.RecordSource = "qry_" & suffix
    Refresh_Captions
End Sub
