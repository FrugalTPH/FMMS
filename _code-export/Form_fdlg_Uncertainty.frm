VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_fdlg_Uncertainty"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private lng_dir As Long
Private s_title As String
Private s_uncertainty As Long
Public w As Long
Public h As Long


Private Sub btn_Cancel_Click()
    DoCmd.Close acForm, Me.name, acSaveNo
End Sub

Private Sub btn_Ok_Click()
    Form_fsub_Memoranda.Controls("uncertainty") = Nz(txt_R, 0)
    DoCmd.Close acForm, Me.name, acSaveNo
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then btn_Cancel_Click
End Sub

Private Sub Form_Load()
    Set_FormSize Me
End Sub

Private Sub Form_Open(Cancel As Integer)
    
    Set_FormIcon Me, "memoranda"
    
    With Form_fsub_Memoranda
        lng_dir = .dir
        s_title = Nz(.Title, vbNullString)
        s_uncertainty = Nz(.uncertainty, 0)
    End With
    Caption = "Uncertainty   ~   M" & lng_dir & " / No name given"
    If s_title <> vbNullString Then Caption = "Uncertainty   ~   M" & lng_dir & " / " & s_title
    
    txt_R = s_uncertainty
    
    Calc_Type s_uncertainty
    Calc_CL s_uncertainty
    
    Refresh_Controls
    
    With New cls_Permissor
        .Init Form_fsub_Memoranda
        btn_Ok.Enabled = .Can_EditUncertainty
    End With
    
End Sub

Private Sub Form_Resize()
    Limit_FormSize Me
End Sub

Private Sub grp_Consequence_AfterUpdate()
    Calc_R
    Refresh_Controls
End Sub

Private Sub grp_Likelihood_AfterUpdate()
    Calc_R
    Refresh_Controls
End Sub

Private Sub grp_Type_AfterUpdate()
    Calc_R
    Refresh_Controls
End Sub

' ------- '
' PRIVATE '
' ------- '

Private Sub Calc_CL(lng_R As Long)
    Select Case Abs(lng_R)
        Case 100:   grp_Consequence = 4: grp_Likelihood = 4
        Case 94:    grp_Consequence = 4: grp_Likelihood = 3
        Case 89:    grp_Consequence = 4: grp_Likelihood = 2
        Case 86:    grp_Consequence = 4: grp_Likelihood = 1
        Case 85:    grp_Consequence = 4: grp_Likelihood = 0
        Case 83:    grp_Consequence = 3: grp_Likelihood = 4
        Case 75:    grp_Consequence = 3: grp_Likelihood = 3
        Case 69:    grp_Consequence = 3: grp_Likelihood = 2
        Case 68:    grp_Consequence = 2: grp_Likelihood = 4
        Case 65:    grp_Consequence = 3: grp_Likelihood = 1
        Case 63:    grp_Consequence = 3: grp_Likelihood = 0
        Case 58:    grp_Consequence = 2: grp_Likelihood = 3
        Case 57:    grp_Consequence = 1: grp_Likelihood = 4
        Case 54:    grp_Consequence = 0: grp_Likelihood = 4
        Case 50:    grp_Consequence = 2: grp_Likelihood = 2
        Case 45:    grp_Consequence = 1: grp_Likelihood = 3
        Case 44:    grp_Consequence = 2: grp_Likelihood = 1
        Case 42:    grp_Consequence = 2: grp_Likelihood = 0
        Case 41:    grp_Consequence = 0: grp_Likelihood = 3
        Case 34:    grp_Consequence = 1: grp_Likelihood = 2
        Case 29:    grp_Consequence = 0: grp_Likelihood = 2
        Case 25:    grp_Consequence = 1: grp_Likelihood = 1
        Case 21:    grp_Consequence = 1: grp_Likelihood = 0
        Case 17:    grp_Consequence = 0: grp_Likelihood = 1
        Case 11:    grp_Consequence = 0: grp_Likelihood = 0
        Case Else:  grp_Consequence = 0: grp_Likelihood = 0
    End Select
End Sub

Private Sub Calc_R()
    Dim C As Double: C = CDbl(grp_Consequence): If C = 0 Then C = 0.5
    Dim L As Double: L = CDbl(grp_Likelihood)
    Dim r As Double: r = 100 * (Sqr((L * L) + (2.5 * C * C)) / Sqr(56))
    txt_R = grp_Type * CLng(r)
End Sub

Private Sub Calc_Type(lng_R As Long)
    grp_Type = 0
    If lng_R < 0 Then grp_Type = -1
    If lng_R > 0 Then grp_Type = 1
End Sub

Private Sub Refresh_Controls()
    grp_Consequence.Enabled = grp_Type <> 0
    grp_Likelihood.Enabled = grp_Type <> 0
    Select Case grp_Type
        Case -1:
            lbl_UR.Caption = "% RISK"
            FormDetail.BackColor = Pastel_1.Red
            FormFooter.BackColor = Pastel_1.Red
            lbl_Type.BackColor = Pastel_1.Red
            lbl_C.BackColor = Pastel_1.Red
            lbl_L.BackColor = Pastel_1.Red
        Case 1:
            lbl_UR.Caption = "% OPPORTUNITY"
            FormDetail.BackColor = Pastel_1.Blue
            FormFooter.BackColor = Pastel_1.Blue
            lbl_Type.BackColor = Pastel_1.Blue
            lbl_C.BackColor = Pastel_1.Blue
            lbl_L.BackColor = Pastel_1.Blue
        Case Else:
            lbl_UR.Caption = "% CLOSED"
            FormDetail.BackColor = Solid.White
            FormFooter.BackColor = Solid.White
            lbl_Type.BackColor = Solid.White
            lbl_C.BackColor = Solid.White
            lbl_L.BackColor = Solid.White
    End Select
    'Debug.Print "w=" & Me.InsideWidth
    'Debug.Print "h=" & Me.InsideHeight
End Sub

' ------ '
' PUBLIC '
' ------ '
