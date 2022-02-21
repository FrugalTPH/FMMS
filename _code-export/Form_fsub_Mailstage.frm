VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_fsub_Mailstage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const prefix As String = "E"
Private Const suffix As String = "_Mailstage"
Private ShiftTest As Integer


Private Sub classCode_DblClick(Cancel As Integer)
    Manual_ClassifyFile Me
End Sub

Private Sub classCode_Enter()
    Auto_ClassifyFile Me
End Sub

Private Sub classCode_GotFocus()
    classCode.Requery
End Sub

Private Sub effectiveOrg_Enter()
    With Me
        If Nz(!effectiveOrg, 0) > 0 Then Exit Sub
        Dim org As Long: org = Auto_EmailOrg(!from)
        If org > 0 Then !effectiveOrg = org
    End With
End Sub

Private Sub effectiveOrg_GotFocus()
    effectiveOrg.Requery
End Sub

Private Sub effectiveOrg_NotInList(newData As String, Response As Integer)
    Response = Organisation_CreateNotInList(effectiveOrg, newData)
End Sub

Private Sub Form_AfterDelConfirm(status As Integer)
    Form_Current
End Sub

Private Sub Form_ApplyFilter(Cancel As Integer, ApplyType As Integer)
    Me.AllowAdditions = ApplyType <> 0
End Sub

Private Sub Form_BeforeDelConfirm(Cancel As Integer, Response As Integer)
    Me.AllowAdditions = True
End Sub

Private Sub Form_Current()
    Form_frm_Mailstage.Refresh_Controls VisibleRecords_Count > 0
    Form_frm_Mailstage.Refresh_Captions
    Me.AllowAdditions = False
End Sub

Private Sub Form_Open(Cancel As Integer)
    Me.RecordSource = "app" & suffix
End Sub

Private Sub ID_Click()
    With Me.RecordsetClone
        .FindFirst "ID=" & Me!ID
        If .NoMatch Then Exit Sub
        If ShiftTest = 1 Then Execute !fdir Else Execute !fdir & !fname & "." & !ftype
    End With
End Sub

Private Sub ID_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ShiftTest = Shift And 1
End Sub

Private Sub mcount_Click()
    Manual_Match Me
End Sub

Private Sub secCode_DblClick(Cancel As Integer)
    Manual_ClassifyFile Me
End Sub

Private Sub secCode_Enter()
    Auto_ClassifyFile Me
End Sub

Private Sub secCode_GotFocus()
    secCode.Requery
End Sub

Private Sub selected_AfterUpdate()
    Me.Dirty = False
    Form_frm_Mailstage.Refresh_Controls VisibleRecords_Count > 0
End Sub

Private Sub statusCode_DblClick(Cancel As Integer)
    Manual_ClassifyFile Me
End Sub

Private Sub statusCode_Enter()
    Auto_ClassifyFile Me
End Sub

Private Sub statusCode_GotFocus()
    statusCode.Requery
End Sub

Private Sub typeCode_DblClick(Cancel As Integer)
    Manual_ClassifyFile Me
End Sub

Private Sub typeCode_Enter()
    Auto_ClassifyFile Me
End Sub

Private Sub typeCode_GotFocus()
    TypeCode.Requery
End Sub

' ------- '
' PRIVATE '
' ------- '

' ------ '
' PUBLIC '
' ------ '

Public Sub CheckMatches()
    Dim rst As Recordset: Set rst = Me.RecordsetClone
    If (rst.EOF And rst.BOF) Then GoTo MyExit
    rst.MoveFirst
    Dim matches As String
    Dim mcount As Long
    Do Until rst.EOF = True
        matches = g_Model.Email_Matches(rst!hash, rst!hashq)
        mcount = Nz(ArrayLen(Split(matches, ",")), 0)
        rst.Edit
        rst!matches = matches
        rst!mcount = mcount
        rst.Update
        rst.MoveNext
    Loop
MyExit:
    rst.Close
    Set rst = Nothing
End Sub

Public Property Get Prefix_() As String
    Prefix_ = prefix
End Property

Public Sub SelectedRecords_ApplyPresets()
    
    Dim rst As Recordset: Set rst = VisibleRecords
    If (rst.EOF And rst.BOF) Then GoTo MyExit
    rst.MoveFirst
    
    With Me.Parent
        Dim effectiveOrg As Long: effectiveOrg = .effectiveOrg
        Dim TypeCode As String: TypeCode = .TypeCode
        Dim secCode As String: secCode = .secCode
        Dim classCode As String: classCode = .classCode
        Dim statusCode As String: statusCode = .statusCode
    End With
    
    Do Until rst.EOF = True
        If rst!selected Then
            rst.Edit
            rst!effectiveOrg = effectiveOrg
            rst!TypeCode = TypeCode
            rst!secCode = secCode
            rst!classCode = classCode
            rst!statusCode = statusCode
            rst.Update
        End If
        rst.MoveNext
    Loop
MyExit:
    rst.Close
    Set rst = Nothing
End Sub

Public Sub SelectedRecords_SetString(k As String, v As String)
    Dim rst As Recordset: Set rst = VisibleRecords
    If (rst.EOF And rst.BOF) Then GoTo MyExit
    rst.MoveFirst
    Do Until rst.EOF = True
        If rst!selected Then
            rst.Edit
            rst.Fields(k).value = v
            rst.Update
        End If
        rst.MoveNext
    Loop
MyExit:
    rst.Close
    Set rst = Nothing
End Sub

Public Sub SelectedRecords_SetLong(k As String, v As Long)
    Dim rst As Recordset: Set rst = VisibleRecords
    If (rst.EOF And rst.BOF) Then GoTo MyExit
    rst.MoveFirst
    Do Until rst.EOF = True
        If rst!selected Then
            rst.Edit
            rst.Fields(k).value = v
            rst.Update
        End If
        rst.MoveNext
    Loop
MyExit:
    rst.Close
    Set rst = Nothing
End Sub

Public Property Get Suffix_() As String
    Suffix_ = suffix
End Property

Public Function VisibleRecords() As Recordset
    Dim rst As Recordset: Set rst = Me.RecordsetClone
    If Me.FilterOn Then rst.Filter = Me.Filter
    Set VisibleRecords = rst.OpenRecordset
End Function

Public Function VisibleRecords_Checked() As Boolean
    VisibleRecords_Checked = False
    Dim rst As Recordset: Set rst = VisibleRecords
    If (rst.EOF And rst.BOF) Then GoTo MyExit
    rst.MoveFirst
    Do Until rst.EOF = True Or VisibleRecords_Checked = True
        VisibleRecords_Checked = Nz(rst!selected, False)
        rst.MoveNext
    Loop
MyExit:
    rst.Close
    Set rst = Nothing
End Function

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

Public Sub VisibleRecords_Selected(bln As Boolean)
    Dim rst As Recordset: Set rst = VisibleRecords
    If (rst.EOF And rst.BOF) Then GoTo MyExit
    rst.MoveFirst
    Do Until rst.EOF = True
        rst.Edit
        rst!selected = bln
        rst.Update
        rst.MoveNext
    Loop
MyExit:
    rst.Close
    Set rst = Nothing
End Sub

Public Function VisibleRecords_UnChecked() As Boolean
    VisibleRecords_UnChecked = False
    Dim rst As Recordset: Set rst = VisibleRecords
    If (rst.EOF And rst.BOF) Then GoTo MyExit
    rst.MoveFirst
    Do Until rst.EOF = True Or VisibleRecords_UnChecked = True
        VisibleRecords_UnChecked = Nz(Not rst!selected, False)
        rst.MoveNext
    Loop
MyExit:
    rst.Close
    Set rst = Nothing
End Function
