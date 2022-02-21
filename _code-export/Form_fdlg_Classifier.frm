VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_fdlg_Classifier"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const suffix As String = "Classifier"
Private s_FormName As String
Private s_fieldName As String
Public w As Long
Public h As Long


Private Sub classCode_DblClick(Cancel As Integer)
    SaveAndExit
End Sub

Private Sub ClassName_DblClick(Cancel As Integer)
    SaveAndExit
End Sub

Private Sub Form_DblClick(Cancel As Integer)
    SaveAndExit
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then CancelAndExit
End Sub

Private Sub Form_Open(Cancel As Integer)

    Set_FormIcon Me, LCase$(suffix)
    Set_FormSize Me
    
    Dim args() As String: args = Split(Nz(Me.OpenArgs, vbNullString), ";")
    s_FormName = args(0)
    s_fieldName = args(1)
    Dim strSql As String: strSql = std_Sql.fdlg_Classifier(s_FormName, s_fieldName)
    If strSql = vbNullString Then CancelAndExit
    Me.RecordSource = strSql
    Me.txt_Query = vbNullString
    
    Me.Caption = "Set Document "
    If s_FormName = "fsub_Schemes" Then Me.Caption = "Set Default "
    If s_FormName = "frm_Filestage" Then Me.Caption = "Set Default "
    
    Select Case s_fieldName
        Case "classCode": Me.Caption = Me.Caption & "CLASSIFICATION (dbl-clk):"
        Case "locCode": Me.Caption = Me.Caption & "LOCATION (dbl-clk):"
        Case "origCode": Me.Caption = Me.Caption & "ORIGINATOR (dbl-clk):"
        Case "projectCode": Me.Caption = Me.Caption & "PROJECT (dbl-clk):"
        Case "revCode": Me.Caption = Me.Caption & "REVISION (dbl-clk):"
        Case "roleCode": Me.Caption = Me.Caption & "ROLE (dbl-clk):"
        Case "secCode": Me.Caption = Me.Caption & "SECURITY (dbl-clk):"
        Case "statusCode": Me.Caption = Me.Caption & "STATUS (dbl-clk):"
        Case "sysCode": Me.Caption = Me.Caption & "SYSTEM (dbl-clk):"
        Case "typeCode": Me.Caption = Me.Caption & "DOCUMENT TYPE (dbl-clk):"
        Case "wsCode": Me.Caption = Me.Caption & "WORK-STAGE (dbl-clk):"
        'Case "schemeCode": Me.Caption = Me.Caption & "SCHEME (dbl-clk):"
        Case Else: Me.Caption = "Classifier"
    End Select
    
End Sub

Private Sub txt_Query_Change()
    Me.Requery
    If Me.RecordsetClone.RecordCount > 0 Then
        Me.txt_Query.SelStart = Me.txt_Query.SelLength
    Else
        Me.txt_Query = Left(Me.txt_Query, Len(Me.txt_Query) - 1)
        txt_Query_Change
    End If
End Sub

' ------- '
' PRIVATE '
' ------- '

Private Sub CancelAndExit()
    DoCmd.Close acForm, Me.name, acSaveNo
End Sub

Private Sub SaveAndExit()
    If g_Model.Db_isReadOnly Then Exit Sub
    If Not Me.classCode = vbNullString Then
        Dim frm As Form: Set frm = GetForm(s_FormName)
        frm.Controls(s_fieldName) = Me.classCode
        Set frm = Nothing
    End If
    DoCmd.Close acForm, Me.name, acSaveNo
End Sub

' ------ '
' PUBLIC '
' ------ '

