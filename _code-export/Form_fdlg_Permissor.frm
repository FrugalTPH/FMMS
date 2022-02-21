VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_fdlg_Permissor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private lng_dir As Long
Private s_title As String
Private s_permissions As String
Public w As Long
Public h As Long


Private Sub btn_Cancel_Click()
    DoCmd.Close acForm, Me.name, acSaveNo
End Sub

Private Sub btn_Ok_Click()
    g_Model.Entity_Archive "tbl_People", lng_dir
    g_Model.Entity_Update "tbl_People", lng_dir, "permissions='" & s_permissions & "'"
    DoCmd.Close acForm, Me.name, acSaveNo
    Form_fsub_People.Requery
End Sub

Private Sub chk_Manager_AfterUpdate()
    Refresh_Controls
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then btn_Cancel_Click
End Sub

Private Sub Form_Load()
    Set_FormSize Me
End Sub

Private Sub Form_Open(Cancel As Integer)
    Set_FormIcon Me, "permissions"
    Set_FormSize Me
    lst_Permissions.RowSource = std_Sql.fdlg_Permissor
    With Form_fsub_People
        lng_dir = .dir
        s_title = Nz(.Title, vbNullString)
        s_permissions = Nz(.permissions, vbNullString)
    End With
    Caption = "Permissions   ~   P" & lng_dir & " / No name given"
    If s_title <> vbNullString Then Caption = "Permissions   ~   P" & lng_dir & " / " & s_title
    Init_Controls s_permissions
End Sub

Private Sub Form_Resize()
    Limit_FormSize Me
End Sub

Private Sub lst_Permissions_AfterUpdate()
    Refresh_Controls
End Sub

' ------- '
' PRIVATE '
' ------- '

Private Sub Init_Controls(strCsv As String)
    chk_Manager = InStr(s_permissions, ",*,") > 0
    If Not chk_Manager Then
        Dim lngRow As Long
        With lst_Permissions
            For lngRow = 0 To .ListCount - 1
                .selected(lngRow) = InStr(strCsv, "," & .Column(0, lngRow) & ",") > 0
            Next
        End With
    End If
    lst_Permissions.Enabled = Not chk_Manager
    cur_Permissions.Caption = "N/A"
    If s_permissions <> vbNullString Then cur_Permissions.Caption = s_permissions
End Sub

Private Sub Refresh_Controls()
    s_permissions = vbNullString
    If chk_Manager Then s_permissions = ",*,"
    If Not chk_Manager Then
        Dim vItem As Variant
        For Each vItem In lst_Permissions.ItemsSelected
            s_permissions = s_permissions & "," & lst_Permissions.ItemData(vItem)
        Next
        If s_permissions <> vbNullString Then s_permissions = s_permissions & ","
    End If
    lst_Permissions.Enabled = Not chk_Manager
    cur_Permissions.Caption = "N/A"
    If s_permissions <> vbNullString Then cur_Permissions.Caption = s_permissions
    'Debug.Print "w=" & Me.InsideWidth
    'Debug.Print "h=" & Me.InsideHeight
End Sub

' ------ '
' PUBLIC '
' ------ '

