VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_fdlg_Templater"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private s_FormName As String
Private prefix As String
Private l_dir As Long
Public w As Long
Public h As Long


Private Sub btn_Apply_Click()
    Dim src As Long: src = Nz(sub_Browser.Form!dir, 0)
    If src <= 0 Then GoTo MyExit
    If Not g_Model.Dir_Copy("N", src, prefix, l_dir) Then GoTo MyExit
    With GetForm(s_FormName)
        !TypeCode = cmb_Main
    End With
MyExit:
    DoCmd.Close acForm, Me.name, acSaveNo
End Sub

Private Sub btn_Cancel_Click()
    DoCmd.Close acForm, Me.name, acSaveNo
End Sub

Private Sub cmb_Main_AfterUpdate()
    With sub_Browser.Form
        .Requery
        Refresh_Controls .VisibleRecords_Count > 0
    End With
    Refresh_Captions
End Sub

Private Sub Form_Load()
    Set_FormSize Me
    Refresh_RecordSource
    With sub_Browser.Form
        .Requery
        Refresh_Controls .VisibleRecords_Count > 0
    End With
    Refresh_Captions
End Sub

Private Sub Form_Open(Cancel As Integer)
    If g_Model.Db_isReadOnly Then btn_Cancel_Click
    Set_FormSize Me
    s_FormName = Me.OpenArgs
    Dim frm As Form: Set frm = GetForm(s_FormName)
    With frm
        prefix = .Prefix_
        l_dir = Nz(!dir, 0)
        FormHeader.BackColor = .Parent.FormHeader.BackColor
        Detail.BackColor = .Parent.Detail.BackColor
        FormFooter.BackColor = .Parent.FormFooter.BackColor
        cmb_Main.RowSource = .TypeCode.RowSource
        cmb_Main = !TypeCode
        If prefix = "U" Then
            Set_FormIcon Me, "outputs"
            If Nz(cmb_Main, vbNullString) = vbNullString Then cmb_Main = TryGetSchemeValue("typeCode", prefix)
        Else
            Set_FormIcon Me, "memoranda"
            If Nz(cmb_Main, vbNullString) = vbNullString Then cmb_Main = TryGetClassification("typeCode", prefix)
        End If
    End With
    btn_Apply.Caption = "Copy to " & prefix & l_dir
End Sub

Private Sub Form_Resize()
    Limit_FormSize Me
End Sub

' ------- '
' PRIVATE '
' ------- '

Private Sub Refresh_RecordSource()
On Error Resume Next
    Query_Refresh "qry_Templater", "SELECT * FROM tbl_Inputs WHERE isTemplate=True AND typeCode=[Forms]![fdlg_Templater]![cmb_Main];"
    sub_Browser.Form.RecordSource = "qry_Templater"
End Sub

' ------ '
' PUBLIC '
' ------ '

Private Property Get Form_Title() As String
    Form_Title = cmb_Main.Column(1) & " templates   ~   None available"
    Dim frm As Form: Set frm = sub_Browser.Form
    If frm.Recordset.RecordCount > 0 Then
        Dim dir As Long: dir = Nz(frm!dir, 0)
        Dim Title As String: Title = Nz(frm!Title, vbNullString)
        If dir = 0 Or Title = vbNullString Then Exit Property
        Form_Title = cmb_Main.Column(1) & " templates   ~   N" & dir & "  /  " & mod_StringUtils.StripNonAsciiChars(Title)
    End If
End Property

Public Sub Refresh_Captions()
    Caption = Form_Title
End Sub

Public Sub Refresh_Controls(Optional bln_hasRecords As Boolean = False)
    btn_Apply.Enabled = bln_hasRecords
End Sub
