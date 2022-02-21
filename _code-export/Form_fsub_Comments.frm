VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_fsub_Comments"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub Form_BeforeInsert(Cancel As Integer)
    SetDefaults_Create
End Sub

Private Sub sysStartPerson_GotFocus()
    sysStartPerson.Requery
End Sub


' ------- '
' PRIVATE '
' ------- '

Private Sub SetDefaults_Create()
    Me.sysStartPerson = g_Model.User_Current
    Me.sysGuid = CreateGuid
End Sub

' ------ '
' PUBLIC '
' ------ '

