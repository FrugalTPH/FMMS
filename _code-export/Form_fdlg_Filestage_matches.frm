VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_fdlg_Filestage_matches"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const prefix As String = "N"
Private Const suffix As String = "Inputs"
Private csv_Records As String
Public w As Long
Public h As Long


Private Sub btn_Select_Click()
    SaveAndExit
End Sub

Private Sub classCode_GotFocus()
    classCode.Requery
End Sub

Private Sub Dir_Click()
    DirClick Me
End Sub

Private Sub Form_Open(Cancel As Integer)
    
    Set_FormIcon Me, LCase$(suffix)
    Set_FormSize Me
    
    Dim args As String: args = Me.OpenArgs
    csv_Records = Nz(args, vbNullString)
    Dim strSql As String: strSql = std_Sql.fdlg_Filestage_matches(csv_Records)
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

Private Sub SaveAndExit()
    If Not Me.dir = vbNullString Then
        Dim frm As Form: Set frm = Forms("frm_Filestage").sub_Browser.Form
        frm!target = Me!dir
        frm!Title = Me!Title
        frm!aliases = Me!aliases
        frm!TypeCode = Me!TypeCode
        frm!secCode = Me!secCode
        frm!classCode = Me!classCode
        frm!statusCode = Me!statusCode
        frm!refNo = Me!refNo
        frm!effectiveOrg = Me!effectiveOrg
        Set frm = Nothing
    End If
    DoCmd.Close acForm, Me.name, acSaveNo
End Sub

' ------ '
' PUBLIC '
' ------ '

Public Property Get Prefix_() As String
    Prefix_ = prefix
End Property
