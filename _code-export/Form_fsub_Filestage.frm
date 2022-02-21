VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_fsub_Filestage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const prefix As String = "N"
Private Const suffix As String = "_Filestage"
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

Private Sub effectiveDate_Enter()
    If Nz(Me!effectiveDate, 0) = 0 Then Me!effectiveDate = ToDDMMYYYY(mindate(Me!fDateLastModified, Me!fDateCreated))
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
    If ApplyType = 0 Then
        Me.AllowAdditions = False
    Else
        Me.AllowAdditions = True
    End If
End Sub

Private Sub Form_BeforeDelConfirm(Cancel As Integer, Response As Integer)
    Me.AllowAdditions = True
End Sub

Private Sub Form_Current()
    Form_frm_Filestage.Refresh_Controls VisibleRecords_Count > 0
    Form_frm_Filestage.Refresh_Captions
    Me.AllowAdditions = False
End Sub

Private Sub Form_Open(Cancel As Integer)
    Me.RecordSource = "app" & suffix
End Sub

Private Sub ID_Click()
    Dim trg As String: trg = Me!fullpath
    If ShiftTest = 1 Then trg = Me!fdir
    Execute trg
End Sub

Private Sub ID_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ShiftTest = Shift And 1
End Sub

Private Sub mcount_Click()
    Manual_Match Me
End Sub

Private Sub refNo_AfterUpdate()
    With Me
        If Not IsNull(!refNo) And Len(!refNo) > 0 Then !refNo = ValidDocRefNo(!refNo)
        !matches = g_Model.File_Matches(!hash, !hashq, Nz(!refNo, vbNullString))
        !mcount = Nz(ArrayLen(Split(!matches, ",")), 0)
    End With
End Sub

Private Sub revNo_AfterUpdate()
    If Not IsNull(revNo) And Len(revNo) > 0 Then revNo = ValidDocRevNo(revNo)
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
    Form_frm_Filestage.Refresh_Controls VisibleRecords_Count > 0
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

Private Sub target_AfterUpdate()

    If IsEmpty(target) Or IsNull(target) Then Exit Sub
    
    Dim rst As Recordset
    
    If IsNumeric(target) Then
        Set rst = dbLocal.OpenRecordset("tbl_Inputs", dbOpenDynaset)
        rst.FindFirst "dir=" & target
        If rst.NoMatch Then GoTo MyExit
    Else
        Set rst = Me.Form.RecordsetClone
        rst.FindFirst "target='" & target & "'"
        If rst.NoMatch Then GoTo MyExit
        revNo = rst!revNo
        effectiveDate = rst!effectiveDate
    End If
    
    Title = rst!Title
    aliases = rst!aliases
    TypeCode = rst!TypeCode
    secCode = rst!secCode
    classCode = rst!classCode
    statusCode = rst!statusCode
    refNo = rst!refNo
    effectiveOrg = rst!effectiveOrg

MyExit:
    rst.Close
    Set rst = Nothing
End Sub

Private Sub target_DblClick(Cancel As Integer)
    Dim rst As Recordset: Set rst = Me.RecordsetClone
    rst.MoveLast
    rst.MoveFirst
    If rst.RecordCount <= 0 Then GoTo MyExit
    If MsgBox("Click OK to apply this item's presets to all others of the same TARGET?", vbOKCancel + vbQuestion, "Normalise") <> vbOK Then GoTo MyExit
    Do Until rst.EOF = True
        If rst!target = Me!target Then
            rst.Edit
            rst!Title = Me!Title
            rst!refNo = Me!refNo
            rst!revNo = Me!revNo
            rst!aliases = Me!aliases
            rst!effectiveDate = Me!effectiveDate
            rst!effectiveOrg = Me!effectiveOrg
            rst!TypeCode = Me!TypeCode
            rst!secCode = Me!secCode
            rst!classCode = Me!classCode
            rst!statusCode = Me!statusCode
            rst.Update
        End If
        rst.MoveNext
    Loop
    
MyExit:
    rst.Close
    Set rst = Nothing
End Sub

Private Sub target_GotFocus()
    target.Requery
End Sub

Private Sub target_KeyPress(KeyAscii As Integer)
    If KeyAscii > 47 And KeyAscii < 58 Then Exit Sub
    If KeyAscii > 64 And KeyAscii < 91 Then Exit Sub
    If KeyAscii > 96 And KeyAscii < 123 Then Exit Sub
    If KeyAscii = 8 Then Exit Sub
    If KeyAscii = 127 Then Exit Sub
    KeyAscii = 0
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
        matches = g_Model.File_Matches(rst!hash, rst!hashq, Nz(rst!refNo, vbNullString))
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
