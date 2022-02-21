VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_fdlg_NewModel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private b_ScriptIsSelected As Boolean


Private Sub btn_Cancel_Click()
    DoCmd.Close acForm, Me.name, acSaveNo
End Sub

Private Sub btn_Ok_Click()
    Dim src As String: src = g_App.App_Scripts & Me.lbo_Scripts
    Dim trg As String: trg = g_App.App_Models & lbl_ModelName.Caption
    g_Model.Db_Create src, trg
    DoCmd.Close acForm, Me.name, acSaveNo
    
    Form_frm_Home.Model_Load trg
    
End Sub

Private Sub Form_Load()
    Set_FormIcon Me, "frugal"
    RefreshListBox
    RefreshButtons
    txt_ModelName = vbNullString
    lbl_ModelName.Caption = vbNullString
End Sub

Private Sub lbo_Scripts_AfterUpdate()
    RefreshButtons
End Sub

Private Sub txt_ModelName_AfterUpdate()
    lbl_ModelName.Caption = mod_StringUtils.KebabCase(txt_ModelName.Text)
    If lbl_ModelName.Caption <> vbNullString Then lbl_ModelName.Caption = lbl_ModelName.Caption & ".accdb"
    RefreshButtons
End Sub

Private Sub txt_ModelName_KeyUp(KeyCode As Integer, Shift As Integer)
    lbl_ModelName.Caption = mod_StringUtils.KebabCase(txt_ModelName.Text)
    If lbl_ModelName.Caption <> vbNullString Then lbl_ModelName.Caption = lbl_ModelName.Caption & ".accdb"
End Sub

' ------- '
' PRIVATE '
' ------- '

Private Sub RefreshListBox()
    Me.lbo_Scripts.value = Nothing
    Dim strFile As String: strFile = dir(g_App.App_Scripts & "*.accdb", vbNormal)
    Do While Len(strFile) > 0
        Me.lbo_Scripts.AddItem strFile
        strFile = dir()
    Loop
    lbo_Scripts = lbo_Scripts.ItemData(0)
End Sub

Private Sub RefreshButtons()
    Me.btn_Ok.Enabled = False
    Me.lbl_Script.Visible = False
    
    Me.lbl_Script.Caption = Nz(Me.lbo_Scripts, "Error - no scripts found!")
    
    If IsNull(Me.lbo_Scripts) Or Me.lbo_Scripts = vbNullString Then
        Me.lbl_Script.Visible = True
        Exit Sub
    End If
    
    If lbl_ModelName.Caption = vbNullString Then
        Exit Sub
    End If
    
    Me.btn_Ok.Enabled = True
    
End Sub

' ------ '
' PUBLIC '
' ------ '
