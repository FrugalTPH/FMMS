VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frm_Schemes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const prefix As String = "S"
Private Const suffix As String = "Schemes"
Private ShiftTest As Integer
Private curParents As String
Private curChildren As String
Private sql_SchemeAttch As String
Private sql_ParentAttch As String
Private sql_ChildAttch As String
Private sql_SchemeOutputs As String
Private sql_ParentOutputs As String
Private sql_ChildOutputs As String
Private bln_AllowForcedClose As Boolean
Public w As Long
Public h As Long


Private Sub btn_Emails_Click()
    If ShiftTest = 1 Then
        DoCmd.OpenForm "frm_Mailstage"
    Else
        ToggleForm "frm_Emails"
    End If
End Sub

Private Sub btn_Emails_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ShiftTest = Shift And 1
End Sub

Private Sub btn_Definitions_Click()
    ToggleForm "frm_Definitions"
End Sub

Private Sub btn_Memoranda_Click()
    ToggleForm "frm_Memoranda"
End Sub

Private Sub btn_Inputs_Click()
    If ShiftTest = 1 Then
        DoCmd.OpenForm "frm_Filestage"
    Else
        ToggleForm "frm_Inputs"
    End If
End Sub

Private Sub btn_Inputs_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ShiftTest = Shift And 1
End Sub

Private Sub btn_Organisations_Click()
    ToggleForm "frm_Organisations"
End Sub

Private Sub btn_Outputs_Click()
    ToggleForm "frm_Outputs"
End Sub

Private Sub btn_People_Click()
    ToggleForm "frm_People"
End Sub

Private Sub btn_Templates_Click()
    If ShiftTest = 1 Then
        g_App.Templates_Open
    Else
        g_Model.Templates_Open
    End If
End Sub

Private Sub btn_Templates_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ShiftTest = Shift And 1
End Sub

Private Sub chk_Deleted_AfterUpdate()
    Refresh_RecordSource
End Sub

Private Sub cmb_Main_AfterUpdate()
    Refresh_RecordSource
End Sub

Private Sub cmb_Main_KeyDown(KeyCode As Integer, Shift As Integer)
    KeyCode = 0
End Sub

Private Sub Form_Activate()
    If g_Model.Definition_Refresh Then
        cmb_Main.Requery
        cmb_Main = cmb_Main.Column(1, 0)
        sub_Browser.Form.Requery
    End If

End Sub

Private Sub Form_Close()
    DoEvents
    mod_ViewUtils.CloseOpenForms Me.name
    If g_App.App_Mode = LocalSnapshot Then
        g_Model.Terminate
        g_App.Terminate
    Else
        DoCmd.OpenForm "frm_Home"
    End If
End Sub

Private Sub Form_Load()
    
    If BootBlockedUser Then Exit Sub
    
    Set_FormIcon Me, LCase$(suffix)
    Set_FormSize Me
    
    FormHeader.BackColor = Pastel_0.Gray
    Detail.BackColor = Pastel_0.Gray
    FormFooter.BackColor = Pastel_0.Gray
    
    cmb_Main.RowSource = std_Sql.cmb_Main_S
    cmb_Main = cmb_Main.Column(1, 0)
    
    btn_Definitions.BackColor = Solid.Black
    btn_Definitions.Properties("hovercolor") = Solid.Black
    btn_Definitions.Properties("pressedcolor") = Solid.Black
    btn_Definitions.BackStyle = 0

    btn_People.BackColor = Pastel_1.Yellow
    btn_People.Properties("hovercolor") = Pastel_0.Yellow
    btn_People.Properties("pressedcolor") = Pastel_0.Yellow
    btn_People.BackStyle = 0
    
    btn_Organisations.BackColor = Pastel_1.Indigo
    btn_Organisations.Properties("hovercolor") = Pastel_0.Indigo
    btn_Organisations.Properties("pressedcolor") = Pastel_0.Indigo
    btn_Organisations.BackStyle = 0
    
    btn_Emails.BackColor = Pastel_1.Grape
    btn_Emails.Properties("hovercolor") = Pastel_0.Grape
    btn_Emails.Properties("pressedcolor") = Pastel_0.Grape
    btn_Emails.BackStyle = 0
    
    btn_Memoranda.BackColor = Pastel_1.Red
    btn_Memoranda.Properties("hovercolor") = Pastel_0.Red
    btn_Memoranda.Properties("pressedcolor") = Pastel_0.Red
    btn_Memoranda.BackStyle = 0
    
    btn_Inputs.BackColor = Pastel_1.Green
    btn_Inputs.Properties("hovercolor") = Pastel_0.Green
    btn_Inputs.Properties("pressedcolor") = Pastel_0.Green
    btn_Inputs.BackStyle = 0
    
    btn_Outputs.BackColor = Pastel_1.Blue
    btn_Outputs.Properties("hovercolor") = Pastel_0.Blue
    btn_Outputs.Properties("pressedcolor") = Pastel_0.Blue
    btn_Outputs.BackStyle = 0
    
    Refresh_FsRoot
    Refresh_RecordSource

    Set_FormPermissions Me.Form, Form_fsub_Schemes
    
End Sub

Private Sub Form_Resize()
    Limit_FormSize Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If bln_AllowForcedClose Then Exit Sub
    Cancel = MsgBox("Ok to close the current model?", vbOKCancel, "Close Model") <> vbOK
End Sub

' ------- '
' PRIVATE '
' ------- '

Private Function BootBlockedUser() As Boolean
    bln_AllowForcedClose = False
    BootBlockedUser = g_Model.User_IsBlocked
    If BootBlockedUser Then
        MsgBox "You are currently blocked from accessing the selected model.", vbInformation, "Access Denied!"
        ForceClose
    End If
End Function

Private Property Get Form_Title() As String
    Form_Title = "Schemes   ~   No selection"
    Dim frm As Form: Set frm = sub_Browser.Form
    If frm.Recordset.RecordCount > 0 Then
        Dim dir As Long: dir = Nz(frm!dir, 0)
        Dim Title As String: Title = Nz(frm!Title, vbNullString)
        Dim wbs As String: wbs = Nz(frm!wbs, vbNullString)
        If dir = 0 Or Title = vbNullString Then Exit Property
        Form_Title = "Schemes   ~   " & prefix & dir & "  /  " & Nz(wbs + ". ", vbNullString) & mod_StringUtils.StripNonAsciiChars(Title)
    End If
End Property

Private Sub Refresh_Captions()
    Caption = Form_Title
    txt_Parents.Caption = curParents
    txt_Children.Caption = curChildren
End Sub

Private Sub Refresh_RecordSource()
    Dim strSql As String
    If chk_Deleted Then
        strSql = std_Sql.Schemes_old
    Else
        strSql = std_Sql.Schemes_current
    End If
    Query_Refresh "qry_" & suffix, strSql
    sub_Browser.Form.RecordSource = "qry_" & suffix
    Set_Criteria
End Sub

' ------ '
' PUBLIC '
' ------ '

Public Property Get CurrentScheme() As Long
On Error Resume Next
    CurrentScheme = 0
    CurrentScheme = Nz(sub_Browser.Form!dir, 0)
End Property

Public Sub ForceClose()
    bln_AllowForcedClose = True
    DoCmd.Close acForm, Me.name, acSaveNo
End Sub

Public Function Get_Criteria(criteria As S_CriteriaType) As String
    Get_Criteria = vbNullString
    Select Case criteria
        Case Csv_Children: Get_Criteria = curChildren
        Case Csv_Parents: Get_Criteria = curParents
        Case Csv_Scheme: Get_Criteria = "," & CurrentScheme & ","
        Case Sql_Atta_Child: Get_Criteria = sql_ChildAttch
        Case Sql_Atta_Parent: Get_Criteria = sql_ParentAttch
        Case Sql_Atta_Scheme: Get_Criteria = sql_SchemeAttch
        Case Sql_Outp_Child: Get_Criteria = sql_ChildOutputs
        Case Sql_Outp_Parent: Get_Criteria = sql_ParentOutputs
        Case Sql_Outp_Scheme: Get_Criteria = sql_SchemeOutputs
    End Select
End Function

Public Property Get Prefix_() As String
    Prefix_ = prefix
End Property

Public Sub Refresh_FsRoot()
    Dim fsroot As String: fsroot = g_Model.Fs_Root
    If fsroot = vbNullString Then
        txt_FsPath.Caption = "Root: Not Connected"
    Else
        txt_FsPath.Caption = "Root: " & fsroot
    End If
End Sub

Public Sub Set_Criteria()
    Dim curScheme As Long: curScheme = CurrentScheme
    curParents = g_Model.Scheme_Parents(curScheme)
    curChildren = g_Model.Scheme_Children(curScheme)
    sql_SchemeAttch = std_Sql.Where_AttachedToAlike(CStr(curScheme))
    sql_ParentAttch = std_Sql.Where_AttachedToAlike(curParents)
    sql_ChildAttch = std_Sql.Where_AttachedToAlike(curChildren)
    sql_SchemeOutputs = std_Sql.Where_SchemeEquals(CStr(curScheme))
    sql_ParentOutputs = std_Sql.Where_SchemeEquals(curParents)
    sql_ChildOutputs = std_Sql.Where_SchemeEquals(curChildren)
    Refresh_Captions
End Sub

