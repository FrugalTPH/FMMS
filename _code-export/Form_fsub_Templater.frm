VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_fsub_Templater"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const prefix As String = "N"
Private Const suffix As String = "Inputs"


Private Sub classCode_GotFocus()
    classCode.Requery
End Sub

Private Sub Dir_Click()
    DirClick Me
End Sub

Private Sub dir_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = acRightButton And Nz(Me!dir, 0) > 0 Then
        Refresh_ContextMenu
        CommandBars("menu_" & suffix).ShowPopup
        DoCmd.CancelEvent
    End If
End Sub

Private Sub effectiveOrg_GotFocus()
    effectiveOrg.Requery
End Sub

Private Sub Form_Current()
    With Form_fdlg_Templater
        .Refresh_Captions
        .Refresh_Controls VisibleRecords_Count > 0
    End With
End Sub

Private Sub Form_Open(Cancel As Integer)
    Me.RecordSource = "qry_Templater"
End Sub

Private Sub secCode_GotFocus()
    secCode.Requery
End Sub

Private Sub statusCode_GotFocus()
    statusCode.Requery
End Sub

Private Sub sysStartPerson_GotFocus()
    sysStartPerson.Requery
End Sub

Private Sub typeCode_GotFocus()
    TypeCode.Requery
End Sub


' ------- '
' PRIVATE '
' ------- '

Private Sub Refresh_ContextMenu()
    On Error Resume Next
        CommandBars("menu_" & suffix).Delete
    On Error GoTo 0
    Dim cbar As CommandBar: Set cbar = CommandBars.Add(name:="menu_" & suffix, Position:=msoBarPopup)
    With New cls_Permissor
        .Init Me.Form
        AddCustomButton cbar, "Comments", "Input_Comments", 0
        If .Can_ReadFs Then
            AddCustomButton cbar, "Work in Progress", "Input_OpenWIP", 0
            AddCustomButton cbar, "Latest Snapshot", "Input_OpenLatestSnapshot", 0
            AddCustomButton cbar, "All Snapshots", "Input_OpenSnapshots", 0
        End If
    End With
End Sub

Private Function VisibleRecords() As Recordset
    Dim rst As Recordset: Set rst = Me.RecordsetClone
    If Me.FilterOn Then rst.Filter = Me.Filter
    Set VisibleRecords = rst.OpenRecordset
End Function

' ------ '
' PUBLIC '
' ------ '

Public Property Get Prefix_() As String
    Prefix_ = prefix
End Property

Public Function VisibleRecords_Count() As Long
    VisibleRecords_Count = 0
    Dim rst As Recordset: Set rst = VisibleRecords
On Error GoTo MyExit
    If (rst.EOF And rst.BOF) Then GoTo MyExit
    rst.MoveLast
    rst.MoveFirst
    VisibleRecords_Count = rst.RecordCount
MyExit:
    rst.Close
    Set rst = Nothing
End Function
