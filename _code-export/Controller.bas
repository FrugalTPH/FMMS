Attribute VB_Name = "Controller"
Option Explicit

Public Const sfx_old As String = "_old"
Public Const sfx_cache As String = "_cache"
Public Const vbBackSlash As String = "\"
Public Const vbDash As String = "-"
Public Const vbMaxDate As Date = #12/31/9999 11:59:59 PM#
Public Const vbMinDate As Date = #1/1/1970#
Public Const vbNullUser As String = "guest"
Public Const vbSpace As String = " "

'Public Const color_WIP As Long = 16241531                   ' rgb(123,211,247)
'Public Const color_SHARED As Long = 8385273                 ' rgb(249,242,127)
'Public Const color_PUBLISHED As Long = 4231247              ' rgb(79,144,64)
'Public Const color_ARCHIVED As Long = 1938423               ' rgb(247,147,29)

Public Enum AppMode
    Uninitialized = 0
    LocalSnapshot = 1
    RemoteBackend = 2
End Enum

Public Enum S_CriteriaType
    Csv_Children = 1
    Csv_Parents = 2
    Csv_Scheme = 3
    Sql_Atta_Child = 4
    Sql_Atta_Parent = 5
    Sql_Atta_Scheme = 6
    Sql_Outp_Child = 7
    Sql_Outp_Parent = 8
    Sql_Outp_Scheme = 9
End Enum


Public Sub Auto_Classify(frm As Form)
On Error GoTo abort
    With frm
        Select Case True
            Case Nz(.ActiveControl, vbNullString) <> vbNullString: Exit Sub
            Case !isFrozen: Exit Sub
            Case g_Model.Db_isReadOnly: Exit Sub
        End Select
        Dim Class As String: Class = TryGetClassification(.ActiveControl.name, LCase$(.Prefix_))
        If Nz(Class, vbNullString) <> vbNullString Then .ActiveControl = Class
    End With
abort:
End Sub

Public Sub Auto_ClassifyFile(frm As Form)
On Error GoTo abort
    With frm
        Select Case True
            Case Nz(.ActiveControl, vbNullString) <> vbNullString: Exit Sub
            Case g_Model.Db_isReadOnly: Exit Sub
        End Select
        Dim Class As String: Class = TryGetClassification(.ActiveControl.name, LCase$(.Prefix_))
        If Nz(Class, vbNullString) <> vbNullString Then .ActiveControl = Class
    End With
abort:
End Sub

Public Sub Auto_ClassifyByScheme(frm As Form)
On Error GoTo abort
    With frm
        Select Case True
            Case Nz(.ActiveControl, vbNullString) <> vbNullString: Exit Sub
            Case !isFrozen: Exit Sub
            Case g_Model.Db_isReadOnly: Exit Sub
        End Select
        Dim Class As String: Class = TryGetSchemeValue(.ActiveControl.name, LCase$(.Prefix_), Nz(!scheme, 0))
        If Nz(Class, vbNullString) = vbNullString Then Class = TryGetClassification(.ActiveControl.name, LCase$(.Prefix_))
        If Nz(Class, vbNullString) <> vbNullString Then .ActiveControl = Class
    End With
abort:
End Sub

Public Function Auto_EmailOrg(strEmail As String) As Long
On Error GoTo abort
    Auto_EmailOrg = 0
    Dim domain As String: domain = Split(strEmail, "@")(1)
    Auto_EmailOrg = Nz(ELookup("dir", "quni_Organisations", "email ALike '%" & domain & "'"), 0)
abort:
End Function

Public Sub Auto_PersonByScheme(frm As Form, Optional scheme As Long = 0)
'On Error GoTo abort
    With frm
        Select Case True
            Case Nz(.ActiveControl, 0) > 0: Exit Sub
            Case !isFrozen: Exit Sub
            Case g_Model.Db_isReadOnly: Exit Sub
        End Select
        Dim person As Variant: person = Nz(TryGetSchemeValue(.ActiveControl.name, LCase$(.Prefix_), scheme), 0)
        If person > 0 Then .ActiveControl = person
    End With
abort:
End Sub

Public Sub CommentOn(frm As Form)
On Error GoTo abort
    DoCmd.OpenForm "fdlg_Commenter", , , , , , frm.Prefix_ & frm!dir
abort:
End Sub

Public Sub DirClick(frm As Form)
    If Not g_Model.Fs_isConnected Then
        MsgBox "You must be connected to the filesystem repository in order to open directories", vbInformation, "FsRoot: Not Connected"
        Exit Sub
    End If
    g_Model.Dir_Open frm.Prefix_, Nz(frm!dir, 0)
End Sub

Public Sub Email_Attach()
    With Form_fsub_Emails
        Dim scheme As Long: scheme = Form_frm_Schemes.CurrentScheme
        Dim dir As Long: dir = !dir
        Dim attachedTo As String: attachedTo = Nz(!attachedTo, vbNullString)
        attachedTo = CsvAdd_numSort(attachedTo, scheme, False)
        If g_Model.Entity_ArchiveAndUpdate("tbl_Emails", dir, "attachedTo='" & attachedTo & "'") Then .Requery
    End With
End Sub

Public Sub Email_Comments()
    DoCmd.OpenForm "fdlg_Commenter", , , , , , "E" & Form_fsub_Emails!dir
End Sub

Public Sub Email_Detach()
    With Form_fsub_Emails
        Dim scheme As Long: scheme = Form_frm_Schemes.CurrentScheme
        Dim dir As Long: dir = !dir
        Dim attachedTo As String: attachedTo = Nz(!attachedTo, vbNullString)
        attachedTo = CsvRemove_numSort(attachedTo, scheme, False)
        If g_Model.Entity_ArchiveAndUpdate("tbl_Emails", dir, "attachedTo='" & attachedTo & "'") Then .Requery
    End With
End Sub

Public Sub Email_IsFrozen()
    With Form_fsub_Emails
        Dim dir As Long: dir = !dir
        Dim isFrozen As Boolean: isFrozen = Not !isFrozen
        If g_Model.Entity_ArchiveAndUpdate("tbl_Emails", dir, "isFrozen=" & isFrozen) Then .Requery
    End With
End Sub

Public Sub Email_OpenWIP()
    g_Model.Dir_Open "E", Form_fsub_Emails!dir
End Sub

Public Sub Email_OpenSnapshots()
    g_Model.Snapshot_Browse "E", Form_fsub_Emails!dir
End Sub

Public Sub Email_OpenLatestSnapshot()
    g_Model.Snapshot_Open "E", Form_fsub_Emails!dir, Form_fsub_Emails!sysEndTime
End Sub

Public Sub Email_CreateSnapshot()
    Dim msg As String: msg = InputBox("Please enter a note to explain the purpose of this snapshot:", "Snapshot Note")
    If StrPtr(msg) = 0 Then Exit Sub
    g_Model.Snapshot_Create "E", Form_fsub_Emails!dir, msg, , , Nz(Form_fsub_Emails!sysGuid, vbNullString)
End Sub

Public Sub Input_Attach()
    With Form_fsub_Inputs
        Dim scheme As Long: scheme = Form_frm_Schemes.CurrentScheme
        Dim dir As Long: dir = !dir
        Dim attachedTo As String: attachedTo = Nz(!attachedTo, vbNullString)
        attachedTo = CsvAdd_numSort(attachedTo, scheme, False)
        If g_Model.Entity_ArchiveAndUpdate("tbl_Inputs", dir, "attachedTo='" & attachedTo & "'") Then .Requery
    End With
End Sub

Public Sub Input_Comments()
    DoCmd.OpenForm "fdlg_Commenter", , , , , , "N" & Form_fsub_Inputs!dir
End Sub

Public Sub Input_CreateSnapshot()
    Dim msg As String: msg = InputBox("Please enter a note to explain the purpose of this snapshot:", "Snapshot Note")
    If StrPtr(msg) = 0 Then Exit Sub
    g_Model.Snapshot_Create "N", Form_fsub_Inputs!dir, msg, , , Nz(Form_fsub_Inputs!sysGuid, vbNullString)
End Sub

Public Sub Input_Detach()
    With Form_fsub_Inputs
        Dim scheme As Long: scheme = Form_frm_Schemes.CurrentScheme
        Dim dir As Long: dir = !dir
        Dim attachedTo As String: attachedTo = Nz(!attachedTo, vbNullString)
        attachedTo = CsvRemove_numSort(attachedTo, scheme, False)
        If g_Model.Entity_ArchiveAndUpdate("tbl_Inputs", dir, "attachedTo='" & attachedTo & "'") Then .Requery
    End With
End Sub

Public Sub Input_IsFrozen()
    With Form_fsub_Inputs
        Dim dir As Long: dir = !dir
        Dim isFrozen As Boolean: isFrozen = Not !isFrozen
        If g_Model.Entity_ArchiveAndUpdate("tbl_Inputs", dir, "isFrozen=" & isFrozen) Then .Requery
    End With
End Sub

Public Sub Input_OpenSnapshots()
    g_Model.Snapshot_Browse "N", Form_fsub_Inputs!dir
End Sub

Public Sub Input_OpenLatestSnapshot()
    g_Model.Snapshot_Open "N", Form_fsub_Inputs!dir, Form_fsub_Inputs!sysEndTime
End Sub

Public Sub Input_OpenWIP()
    g_Model.Dir_Open "N", Form_fsub_Inputs!dir
End Sub

Public Sub LinkClick(frm As Form)
    If Nz(frm!dir, 0) <= 0 Or Nz(frm!link, vbNullString) = vbNullString Then Exit Sub
    Dim str As String: str = Format(frm!dir, frm.dir.Format)
    g_Model.Comment_Create str, "_clickLink"
End Sub

Public Sub Manual_Classify(frm As Form)
On Error GoTo abort
    With frm
        Select Case True
            Case Nz(.ActiveControl.name, vbNullString) = vbNullString: Exit Sub
            Case !isFrozen: Exit Sub
            Case g_Model.Db_isReadOnly: Exit Sub
        End Select
        DoCmd.OpenForm "fdlg_Classifier", , , , , , frm.name & ";" & .ActiveControl.name
    End With
abort:
End Sub

Public Sub Manual_ClassifyFile(frm As Form)
On Error GoTo abort
    With frm
        Select Case True
            Case Nz(.ActiveControl.name, vbNullString) = vbNullString: Exit Sub
            Case g_Model.Db_isReadOnly: Exit Sub
        End Select
        DoCmd.OpenForm "fdlg_Classifier", , , , , , frm.name & ";" & .ActiveControl.name
    End With
abort:
End Sub

Public Sub Manual_Match(frm As Form)
On Error GoTo abort
    With frm
        If Nz(.ActiveControl.name, vbNullString) = vbNullString Then Exit Sub
        If Nz(.ActiveControl, 0) <= 0 Then Exit Sub
        Dim frmName As String: frmName = "fdlg" & .Suffix_ & "_matches"
        Dim strMatches As String: strMatches = .matches
        DoCmd.OpenForm frmName, acFormDS, , , , , strMatches                ' acDialog, strMatches
    End With
abort:
End Sub

Public Sub Manual_Template(frm As Form)
On Error GoTo abort
    With frm
        Select Case True
            Case Nz(.ActiveControl.name, vbNullString) = vbNullString: Exit Sub
            Case !isFrozen: Exit Sub
            Case g_Model.Db_isReadOnly: Exit Sub
            Case Not g_Model.Fs_isConnected: Exit Sub
        End Select
        DoCmd.OpenForm "fdlg_Templater", , , , , , frm.name
    End With
abort:
End Sub

Public Sub Memorandum_Attach()
    With Form_fsub_Memoranda
        Dim scheme As Long: scheme = Form_frm_Schemes.CurrentScheme
        Dim dir As Long: dir = !dir
        Dim attachedTo As String: attachedTo = Nz(!attachedTo, vbNullString)
        attachedTo = CsvAdd_numSort(attachedTo, scheme, False)
        If g_Model.Entity_ArchiveAndUpdate("tbl_Memoranda", dir, "attachedTo='" & attachedTo & "'") Then .Requery
    End With
End Sub

Public Sub Memorandum_Comments()
    DoCmd.OpenForm "fdlg_Commenter", , , , , , "M" & Form_fsub_Memoranda!dir
End Sub

Public Sub Memorandum_Detach()
    With Form_fsub_Memoranda
        Dim scheme As Long: scheme = Form_frm_Schemes.CurrentScheme
        Dim dir As Long: dir = !dir
        Dim attachedTo As String: attachedTo = Nz(!attachedTo, vbNullString)
        attachedTo = CsvRemove_numSort(attachedTo, scheme, False)
        If g_Model.Entity_ArchiveAndUpdate("tbl_Memoranda", dir, "attachedTo='" & attachedTo & "'") Then .Requery
    End With
End Sub

Public Sub Memorandum_Duplicate()
    With Form_fsub_Memoranda
        Dim dir As Long: dir = !dir
        If Nz(dir, 0) = 0 Then Exit Sub
        If MsgBox("Ok to create a new Memo using M" & dir & " as its template?", vbOKCancel, "Duplicate M" & dir) <> vbOK Then Exit Sub
        If g_Model.Memorandum_Duplicate(dir) > 0 Then .Requery
    End With
End Sub

Public Sub Memorandum_IsFrozen()
    With Form_fsub_Memoranda
        Dim dir As Long: dir = !dir
        Dim isFrozen As Boolean: isFrozen = Not !isFrozen
        If g_Model.Entity_ArchiveAndUpdate("tbl_Memoranda", dir, "isFrozen=" & isFrozen) Then .Requery
    End With
End Sub

Public Sub Memorandum_OpenWIP()
    g_Model.Dir_Open "M", Form_fsub_Memoranda!dir
End Sub

Public Sub Memorandum_Uncertainty()
    Dim dir As Long: dir = Form_fsub_Memoranda!dir
    DoCmd.OpenForm "fdlg_Uncertainty", , , , , , dir
End Sub

Public Sub Memorandum_OpenSnapshots()
    g_Model.Snapshot_Browse "M", Form_fsub_Memoranda!dir
End Sub

Public Sub Memorandum_OpenLatestSnapshot()
    g_Model.Snapshot_Open "M", Form_fsub_Memoranda!dir, Form_fsub_Memoranda!sysEndTime
End Sub

Public Sub Memorandum_CreateSnapshot()
    Dim msg As String: msg = InputBox("Please enter a note to explain the purpose of this snapshot:", "Snapshot Note")
    If StrPtr(msg) = 0 Then Exit Sub
    g_Model.Snapshot_Create "M", Form_fsub_Memoranda!dir, msg, , Nz(Form_fsub_Memoranda!location, vbNullString), Nz(Form_fsub_Memoranda!sysGuid, vbNullString)
End Sub

Public Sub Organisation_Attach()
    With Form_fsub_Organisations
        Dim scheme As Long: scheme = Form_frm_Schemes.CurrentScheme
        Dim dir As Long: dir = !dir
        Dim attachedTo As String: attachedTo = Nz(!attachedTo, vbNullString)
        attachedTo = CsvAdd_numSort(attachedTo, scheme, False)
        If g_Model.Entity_ArchiveAndUpdate("tbl_Organisations", dir, "attachedTo='" & attachedTo & "'") Then .Requery
    End With
End Sub

Public Sub Organisation_Comments()
    DoCmd.OpenForm "fdlg_Commenter", , , , , , "O" & Form_fsub_Organisations!dir
End Sub

Public Function Organisation_CreateNotInList(ctl As Control, newData As String) As Integer
On Error GoTo abort
    Select Case True
        Case IsNull(ctl.ControlSource): GoTo abort
        Case IsNull(newData): GoTo abort
        Case MsgBox("Click OK to add this organisation as a new entry in the Organisations register?", vbOKCancel + vbQuestion, "Organisation not In List") <> vbOK: GoTo abort
    End Select
    
    ctl.Undo
    Dim newDir As Long: newDir = g_Model.Organisation_Create(newData)
    If newDir <= 0 Then GoTo abort
    ctl.Requery
    ctl = newDir
    Organisation_CreateNotInList = acDataErrAdded
    Exit Function
abort:
    Organisation_CreateNotInList = acDataErrContinue
End Function

Public Sub Organisation_Detach()
    With Form_fsub_Organisations
        Dim scheme As Long: scheme = Form_frm_Schemes.CurrentScheme
        Dim dir As Long: dir = !dir
        Dim attachedTo As String: attachedTo = Nz(!attachedTo, vbNullString)
        attachedTo = CsvRemove_numSort(attachedTo, scheme, False)
        If g_Model.Entity_ArchiveAndUpdate("tbl_Organisations", dir, "attachedTo='" & attachedTo & "'") Then .Requery
    End With
End Sub

Public Sub Organisation_IsFrozen()
    With Form_fsub_Organisations
        Dim dir As Long: dir = !dir
        Dim isFrozen As Boolean: isFrozen = Not !isFrozen
        If g_Model.Entity_ArchiveAndUpdate("tbl_Organisations", dir, "isFrozen=" & isFrozen) Then .Requery
    End With
End Sub

Public Sub Organisation_OpenWIP()
    g_Model.Dir_Open "O", Form_fsub_Organisations!dir
End Sub

Public Sub Organisation_OpenSnapshots()
    g_Model.Snapshot_Browse "O", Form_fsub_Organisations!dir
End Sub

Public Sub Organisation_OpenLatestSnapshot()
    g_Model.Snapshot_Open "O", Form_fsub_Organisations!dir, Form_fsub_Organisations!sysEndTime
End Sub

Public Sub Organisation_CreateSnapshot()
    Dim msg As String: msg = InputBox("Please enter a note to explain the purpose of this snapshot:", "Snapshot Note")
    If StrPtr(msg) = 0 Then Exit Sub
    g_Model.Snapshot_Create "O", Form_fsub_Organisations!dir, msg, , , Nz(Form_fsub_Organisations!sysGuid, vbNullString)
End Sub

Public Sub Organisation_Undelete()
    With Form_fsub_Organisations
        Dim dir As Long: dir = !dir
        g_Model.Entity_Undelete "tbl_Organisations", dir
        .Requery
    End With
End Sub

Public Sub Output_Attach()
    With Form_fsub_Outputs
        Dim scheme As Long: scheme = Form_frm_Schemes.CurrentScheme
        Dim dir As Long: dir = !dir
        Dim attachedTo As String: attachedTo = Nz(!attachedTo, vbNullString)
        attachedTo = CsvAdd_numSort(attachedTo, scheme, False)
        If g_Model.Entity_ArchiveAndUpdate("tbl_Outputs", dir, "attachedTo='" & attachedTo & "'") Then .Requery
    End With
End Sub

Public Sub Output_Comments()
    DoCmd.OpenForm "fdlg_Commenter", , , , , , "U" & Form_fsub_Outputs!dir
End Sub

Public Sub Output_Detach()
    With Form_fsub_Outputs
        Dim scheme As Long: scheme = Form_frm_Schemes.CurrentScheme
        Dim dir As Long: dir = !dir
        Dim attachedTo As String: attachedTo = Nz(!attachedTo, vbNullString)
        attachedTo = CsvRemove_numSort(attachedTo, scheme, False)
        If g_Model.Entity_ArchiveAndUpdate("tbl_Outputs", dir, "attachedTo='" & attachedTo & "'") Then .Requery
    End With
End Sub

Public Sub Output_Duplicate()
    With Form_fsub_Outputs
        Dim dir As Long: dir = !dir
        If Nz(dir, 0) = 0 Then Exit Sub
        If MsgBox("Ok to create a new Output using U" & dir & " as its template?", vbOKCancel, "Duplicate U" & dir) <> vbOK Then Exit Sub
        If g_Model.Output_Duplicate(dir) > 0 Then .Requery
    End With
End Sub

Public Sub Output_IsFrozen()
    With Form_fsub_Outputs
        Dim dir As Long: dir = !dir
        Dim isFrozen As Boolean: isFrozen = Not !isFrozen
        If g_Model.Entity_ArchiveAndUpdate("tbl_Outputs", dir, "isFrozen=" & isFrozen) Then .Requery
    End With
End Sub

Public Sub Output_OpenWIP()
    g_Model.Dir_Open "U", Form_fsub_Outputs!dir
End Sub

Public Sub Output_OpenSnapshots()
    g_Model.Snapshot_Browse "U", Form_fsub_Outputs!dir
End Sub

Public Sub Output_OpenLatestSnapshot()
    g_Model.Snapshot_Open "U", Form_fsub_Outputs!dir, Form_fsub_Outputs!sysEndTime
End Sub

Public Sub Output_Revise()
    Form_fsub_Outputs.Revise_Current
End Sub

Public Sub Output_CreateSnapshot()
    Dim msg As String: msg = InputBox("Please enter a note to explain the purpose of this snapshot:", "Snapshot Note")
    If StrPtr(msg) = 0 Then Exit Sub
    g_Model.Snapshot_Create "U", Form_fsub_Outputs!dir, msg, , , Nz(Form_fsub_Outputs!sysGuid, vbNullString)
End Sub

Public Sub Person_Attach()
    With Form_fsub_People
        Dim scheme As Long: scheme = Form_frm_Schemes.CurrentScheme
        Dim dir As Long: dir = !dir
        Dim attachedTo As String: attachedTo = Nz(!attachedTo, vbNullString)
        attachedTo = CsvAdd_numSort(attachedTo, scheme, False)
        If g_Model.Entity_ArchiveAndUpdate("tbl_People", dir, "attachedTo='" & attachedTo & "'") Then .Requery
    End With
End Sub

Public Sub Person_Comments()
    DoCmd.OpenForm "fdlg_Commenter", , , , , , "P" & Form_fsub_People!dir
End Sub

Public Function Person_CreateNotInList(ctl As Control, newData As String) As Integer
On Error GoTo abort
    Select Case True
        Case IsNull(ctl.ControlSource): GoTo abort
        Case IsNull(newData): GoTo abort
        Case MsgBox("Click OK to add this person as a new entry in the People register?", vbOKCancel + vbQuestion, "Person not In List") <> vbOK: GoTo abort
    End Select
    
    ctl.Undo
    Dim newDir As Long: newDir = g_Model.Person_Create(newData)
    If newDir <= 0 Then GoTo abort
    ctl.Requery
    ctl = newDir
    Person_CreateNotInList = acDataErrAdded
    Exit Function
abort:
    Person_CreateNotInList = acDataErrContinue
End Function

Public Sub Person_Detach()
    With Form_fsub_People
        Dim scheme As Long: scheme = Form_frm_Schemes.CurrentScheme
        Dim dir As Long: dir = !dir
        Dim attachedTo As String: attachedTo = Nz(!attachedTo, vbNullString)
        attachedTo = CsvRemove_numSort(attachedTo, scheme, False)
        If g_Model.Entity_ArchiveAndUpdate("tbl_People", dir, "attachedTo='" & attachedTo & "'") Then .Requery
    End With
End Sub

Public Sub Person_IsFrozen()
    With Form_fsub_People
        Dim dir As Long: dir = !dir
        Dim isFrozen As Boolean: isFrozen = Not !isFrozen
        If g_Model.Entity_ArchiveAndUpdate("tbl_People", dir, "isFrozen=" & isFrozen) Then .Requery
    End With
End Sub

Public Sub Person_OpenWIP()
    g_Model.Dir_Open "P", Form_fsub_People!dir
End Sub

Public Sub Person_OpenSnapshots()
    g_Model.Snapshot_Browse "P", Form_fsub_People!dir
End Sub

Public Sub Person_OpenLatestSnapshot()
    g_Model.Snapshot_Open "P", Form_fsub_People!dir, Form_fsub_People!sysEndTime
End Sub

Public Sub Person_CreateSnapshot()
    Dim msg As String: msg = InputBox("Please enter a note to explain the purpose of this snapshot:", "Snapshot Note")
    If StrPtr(msg) = 0 Then Exit Sub
    g_Model.Snapshot_Create "P", Form_fsub_People!dir, msg, , , Nz(Form_fsub_People!sysGuid, vbNullString)
End Sub

Public Sub Person_Permissions()
    Dim dir As Long: dir = Form_fsub_People!dir
    DoCmd.OpenForm "fdlg_Permissor", , , , , , dir
End Sub

Public Sub Person_Undelete()
    Dim dir As Long: dir = Form_fsub_People!dir
    g_Model.Entity_Undelete "tbl_People", dir
    Form_fsub_People.Requery
End Sub

Public Sub Scheme_AttachedPeople()
    DoCmd.OpenForm "frm_People"
    With Form_frm_People
        .cmb_Main = 2
        .Refresh_RecordSource
        .Refresh_Controls
    End With
End Sub

Public Sub Scheme_Comments()
    DoCmd.OpenForm "fdlg_Commenter", , , , , , "S" & Form_fsub_Schemes!dir
End Sub

Public Sub Scheme_IsFrozen()
    With Form_fsub_Schemes
        Dim dir As Long: dir = !dir
        Dim isFrozen As Boolean: isFrozen = Not !isFrozen
        If g_Model.Entity_ArchiveAndUpdate("tbl_Schemes", dir, "isFrozen=" & isFrozen) Then .Requery
    End With
End Sub

Public Sub Scheme_OpenWIP()
    g_Model.Dir_Open "S", Form_fsub_Schemes!dir
End Sub

Public Sub Scheme_OpenSnapshots()
    g_Model.Snapshot_Browse "S", Form_fsub_Schemes!dir
End Sub

Public Sub Scheme_OpenLatestSnapshot()
    g_Model.Snapshot_Open "S", Form_fsub_Schemes!dir, Form_fsub_Schemes!sysEndTime
End Sub

Public Sub Scheme_CreateSnapshot()
    Dim msg As String: msg = InputBox("Please enter a note to explain the purpose of this snapshot:", "Snapshot Note")
    If StrPtr(msg) = 0 Then Exit Sub
    g_Model.Snapshot_Create "S", Form_fsub_Schemes!dir, msg, , , Nz(Form_fsub_Schemes!sysGuid, vbNullString)
End Sub

Public Sub Scheme_RenumberWbs()
    If MsgBox("Ok to renumber all elements now?", vbOKCancel, "WBS Renumber") <> vbOK Then Exit Sub
    Dim schemeCode As String: schemeCode = Form_frm_Schemes!cmb_Main
    g_Model.Scheme_RefreshWbs schemeCode
    Form_fsub_Schemes.Requery
End Sub

Public Sub Scheme_Undelete()
    Dim dir As Long: dir = Form_fsub_Schemes!dir
    g_Model.Entity_Undelete "tbl_Schemes", dir
    Form_fsub_Schemes.Requery
End Sub

Public Sub Scheme_WordPad()
    DoCmd.OpenForm "fdlg_WordPad", , , , , , "frm_Schemes"
End Sub
