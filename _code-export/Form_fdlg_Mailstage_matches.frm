VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_fdlg_Mailstage_matches"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const prefix As String = "E"
Private Const suffix As String = "Emails"
Private csv_Records As String
Public w As Long
Public h As Long


Private Sub classCode_GotFocus()
    classCode.Requery
End Sub

Private Sub Dir_Click()
    DirClick Me
End Sub

Private Sub effectiveOrg_GotFocus()
    effectiveOrg.Requery
End Sub

Private Sub Form_Open(Cancel As Integer)
    
    Set_FormIcon Me, LCase$(suffix)
    Set_FormSize Me
    
    Dim args As String: args = Me.OpenArgs
    csv_Records = Nz(args, vbNullString)
    Dim strSql As String: strSql = std_Sql.frm_Mailstage_matches(csv_Records)
    If strSql = vbNullString Then GoTo MyExit
    Dim bln_hasRecords As Boolean: bln_hasRecords = ECount("quni_" & suffix, "dir IN(" & csv_Records & ")") > 0
    If bln_hasRecords = 0 Then
        MsgBox "The potential match could not be traced in the records.", vbExclamation, "Error"
        GoTo MyExit
    End If
    Me.RecordSource = strSql
    Exit Sub
    
MyExit:
    CancelAndExit
End Sub

Private Sub Form_Resize()
    Limit_FormSize Me
End Sub

Private Sub secCode_GotFocus()
    secCode.Requery
End Sub

Private Sub statusCode_GotFocus()
    statusCode.Requery
End Sub

Private Sub typeCode_GotFocus()
    TypeCode.Requery
End Sub

' ------- '
' PRIVATE '
' ------- '

Private Sub CancelAndExit()
    DoCmd.Close acForm, Me.name, acSaveNo
End Sub

' ------ '
' PUBLIC '
' ------ '

Public Property Get Prefix_() As String
    Prefix_ = prefix
End Property

