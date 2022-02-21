Attribute VB_Name = "g_Model"
Option Explicit
Private Const kvsName As String = "tbl_Model"
Private Const temporalTables As String = "tbl_Emails,tbl_Definitions,tbl_Inputs,tbl_Memoranda,tbl_Model,tbl_Organisations,tbl_Outputs,tbl_People,tbl_Schemes,tbl_WordPad"
Private Const standardTables As String = "tbl__Comments,tbl__Files"
Private Const cacheTables As String = "tbl_Definitions"
Private Const ssInfo As String = "#ss_info.txt"


Public Sub Initialize()
    If g_App.App_Mode = LocalSnapshot Then Init_Local Else Init_Remote
End Sub

Public Sub Terminate()
    g_App.Model_Current = vbNullString
    If g_App.App_Mode = LocalSnapshot Then Terminate_Local Else Terminate_Remote
End Sub


'''''''''''''
'' PRIVATE ''
'''''''''''''


Private Sub Init_Remote()

    Dim m As String: m = g_App.Model_Current
    If Db_Validate(m) Then
        
        Dim tbl As Variant
        For Each tbl In Split(standardTables, ",")
            Db_LinkTable m, CStr(tbl)
        Next tbl
        For Each tbl In Split(temporalTables, ",")
            Db_LinkTable m, CStr(tbl)
            Db_LinkTable m, CStr(tbl) & sfx_old
        Next tbl
        For Each tbl In Split(cacheTables, ",")
            Db_ImportTable m, CStr(tbl), CStr(tbl) & sfx_cache
        Next tbl
        
        g_App.App_ValueSet "modelDb", m
        
        If Db_ID = vbNullString Then Db_Init m
        
        Fs_Refresh
        
        User_SignIn
        
        If Not Db_isReadOnly Then
            g_App.App_PropertySet "AppTitle", Db_Name & " - FMMS Editor"
            If DCount("dir", "tbl_People", "permissions ALIKE '%*%'") <= 0 Then User_MakeManager
            If User_Current > 0 Then Comment_Create "model", "Open"
        Else
            g_App.App_PropertySet "AppTitle", Db_Name & " - FMMS Viewer"
            MsgBox "You are connected to the master model, but it is currently archived or on-hold.", vbInformation, "Model Viewer (read-only)"
        End If
        
        Application.RefreshTitleBar
        
    Else
        Terminate
    End If

End Sub

Private Sub Terminate_Remote()

    Dim tbl As Variant
    For Each tbl In Split(standardTables, ",")
        Table_Delete CStr(tbl)
    Next tbl
    For Each tbl In Split(temporalTables, ",")
        Table_Delete CStr(tbl)
        Table_Delete CStr(tbl) & sfx_old
    Next tbl
    
    User_SignOut
    
    g_App.App_PropertySet "AppTitle", "FMMS"
    Application.RefreshTitleBar
    
End Sub

Private Sub Init_Local()

    If Db_Validate(CurrentProject.FullName) Then
    
        g_App.App_PropertySet "AppTitle", Db_Name & " - FMMS Snapshot [" & Db_Date & "]"
        MsgBox "You are viewing a dated snapshot / copy of the master model.", vbInformation, "Model Viewer (read-only)"
    
        Fs_Refresh
        Application.RefreshTitleBar
        
    Else
        Terminate
    End If
    
End Sub

Private Sub Terminate_Local()
    g_App.App_PropertySet "AppTitle", "FMMS"
    Application.RefreshTitleBar
End Sub

Private Property Get Db_ID() As String
    Db_ID = Db_ValueGet("id")
End Property

Private Function Db_IsLoaded() As Boolean
    Db_IsLoaded = g_App.App_ValueGet("modelDb") <> vbNullString
End Function

Private Sub Db_LinkTable(connection As String, tbl As String)
    Table_Delete tbl
    DoCmd.TransferDatabase acLink, "Microsoft Access", connection, acTable, tbl, tbl
End Sub

Private Function Dir_path(prefix As String, dir As Long) As String
    Dir_path = Fs_Root & prefix & vbBackSlash & dir & vbBackSlash
End Function

Private Function Snapshot_path(prefix As String, dir As Long, Optional ss As Double) As String
    Snapshot_path = Fs_Root & prefix & vbBackSlash & dir & vbBackSlash & "_ss" & vbBackSlash
    If ss > 0 Then Snapshot_path = Snapshot_path & ss & vbBackSlash
End Function

Public Sub Templates_Open()
    If Not Fs_isConnected Then
        MsgBox "You must be connected to the filesystem repository in order to open model templates.", vbInformation, "FsRoot: Not Connected"
        Exit Sub
    End If
    Execute MkDirTree(Templates_path)
End Sub

Private Function Templates_path() As String
    Templates_path = Fs_Root & "_templates" & vbBackSlash
End Function

Private Function User_Create() As Long

    User_Create = 0
    If Db_isReadOnly Then Exit Function
    
On Error GoTo User_Create_Error
    Dim rst As Recordset: Set rst = dbLocal.OpenRecordset("tbl_People", dbOpenDynaset)
        rst.AddNew
        rst!identity = g_App.User_Identity
        rst!Title = g_App.User_Name
        rst!email = g_App.User_Email
        rst!sysStartPerson = rst!dir
        rst!object = "P" & rst!dir
        rst.Update
        rst.Bookmark = rst.LastModified
       
    MkDirTree Snapshot_path("P", rst!dir)
    
    User_Create = rst!dir
    
On Error GoTo 0
    rst.Close
    Set rst = Nothing
    Exit Function
    
User_Create_Error:

End Function

Private Function User_IsApprover(roleCode As String) As Boolean
    User_IsApprover = False
    If User_Current <= 0 Then Exit Function
    User_IsApprover = InStr(User_Permissions, "," & roleCode & ",") > 0
End Function

Private Function User_IsSignedIn() As Boolean
    User_IsSignedIn = User_Current > 0
End Function

Private Sub User_MakeManager()
    If Db_isReadOnly Then Exit Sub
    Entity_Archive "tbl_People", User_Current
    Entity_Update "tbl_People", User_Current, "permissions=',*,'"
    MsgBox "You have just been made adminstrator of this model.", vbInformation, "New Administrator"
End Sub

Private Property Get User_Permissions() As String
    User_Permissions = vbNullString
    If User_Current <= 0 Then Exit Property
    User_Permissions = Nz(ELookup("permissions", "tbl_People", "dir=" & User_Current), vbNullString)
End Property

Private Function User_SignIn() As Long
    User_SignIn = User_Current
    If Db_isReadOnly Then Exit Function
    If User_SignIn <= 0 Then User_SignIn = User_Create
    If User_SignIn > 0 Then Comment_Create "model", "Open"
End Function

Private Sub User_SignOut()
    
    g_App.App_ValueRemove "modelDb"
    g_App.App_ValueRemove "modelFs"

    If Not Db_isReadOnly Then Comment_Create "model", "Close"
    
End Sub


''''''''''''
'' PUBLIC ''
''''''''''''

Public Function PublishDb()

    Dim src As String: src = CurrentProject.FullName
    Dim trg As String: trg = Fs_Root & "ss_" & Db_Name & ".accdb"
    
    If Not MsgBox(Fs_Root & vbCrLf & "ss_" & Db_Name & ".accdb" & vbCrLf & vbCrLf & "Click OK to publish to this location now?", vbOKCancel + vbQuestion, "Publish Db") = vbOK Then Exit Function
    
    Dim m As String: m = g_App.Model_Current
    
    Dim tbl As Variant
    For Each tbl In Split(standardTables, ",")
        Table_Delete CStr(tbl)
        Db_ImportTable m, CStr(tbl), CStr(tbl)
    Next tbl
    For Each tbl In Split(temporalTables, ",")
        Table_Delete CStr(tbl)
        Table_Delete CStr(tbl) & sfx_old
        Db_ImportTable m, CStr(tbl), CStr(tbl)
        Db_ImportTable m, CStr(tbl), CStr(tbl) & sfx_old
    Next tbl
    
    File_Copy src, trg, False, coOverwrite
    
    For Each tbl In Split(standardTables, ",")
        Table_Delete CStr(tbl)
        Db_LinkTable m, CStr(tbl)
    Next tbl
    For Each tbl In Split(temporalTables, ",")
        Table_Delete CStr(tbl)
        Table_Delete CStr(tbl) & sfx_old
        Db_LinkTable m, CStr(tbl)
        Db_LinkTable m, CStr(tbl) & sfx_old
    Next tbl
    
    Dim db As DAO.Database: Set db = OpenDatabase(trg)
        db.Execute "DELETE * FROM app_Filestage", dbFailOnError
        db.Execute "DELETE * FROM app_Kvs", dbFailOnError
        db.Execute "DELETE * FROM app_Kvs_old", dbFailOnError
        db.Execute "DELETE * FROM app_Mailstage", dbFailOnError
        db.Execute "INSERT INTO tbl_Model (k,v) VALUES ('date',ToYYYYMMDDHHNN)", dbFailOnError
    db.Close
    Set db = Nothing

    If Not MsgBox(Fs_Root & vbCrLf & "ss_" & Db_Name & ".accdb" & vbCrLf & vbCrLf & "Complete!", vbInformation, "Publish Db") = vbOK Then Exit Function

End Function

Public Function PublishSheet()

    Dim src As String: src = CurrentProject.FullName
    Dim trg As String: trg = Fs_Root & "ss_" & Db_Name & ".xlsx"
    
    If Not MsgBox(Fs_Root & vbCrLf & "ss_" & Db_Name & ".xlsx" & vbCrLf & vbCrLf & "Click OK to publish to this location now?", vbOKCancel + vbQuestion, "Publish Sheet") = vbOK Then Exit Function
    
    DoCmd.TransferSpreadsheet acExport, acSpreadsheetTypeExcel12Xml, "INPUTS", trg, True
    
    If Not MsgBox(Fs_Root & vbCrLf & "ss_" & Db_Name & ".xlsx" & vbCrLf & vbCrLf & "Complete!", vbInformation, "Publish Sheet") = vbOK Then Exit Function

End Function

Public Sub Comment_Create(obj As String, strComment As String)
    If Db_isReadOnly Then Exit Sub
    Dim rst As Recordset: Set rst = dbLocal.OpenRecordset("tbl__Comments", dbOpenDynaset)
        rst.AddNew
        rst!sysStartPerson = User_Current
        rst!object = obj
        rst!comment = strComment
        rst.Update
        rst.Bookmark = rst.LastModified
    rst.Close
    Set rst = Nothing
End Sub

Public Sub Db_Create(src As String, trg As String)
    Dim FSO As New FileSystemObject
    If FSO.FileExists(trg) Then
        MsgBox "A model of this name already exists in this location", vbExclamation, "Cannot create model"
        Exit Sub
    End If
    FSO.CopyFile src, trg
    Db_Read trg
End Sub

Public Property Get Db_Date() As String
    Db_Date = Db_ValueGet("date")
End Property

Private Sub Db_ImportTable(connection As String, src As String, trg As String, Optional structOnly As Boolean = False)
    Table_Delete trg
    DoCmd.TransferDatabase acImport, "Microsoft Access", connection, acTable, src, trg, structOnly
End Sub

Private Sub Db_Init(strPath As String)

    Db_ValueSet "id", CreateGuid
    Db_ValueSet "name", GetBaseName(strPath), True
    Db_ValueSet "is-editable", "true", True
    Comment_Create "model", "Create"
    
    Dim ts As String: ts = "#" & SqlDateTime() & "#"
    With dbLocal
        .Execute "UPDATE tbl_Model SET sysStartTime = " & ts
        .Execute "UPDATE tbl_Definitions SET sysStartTime = " & ts
        .Execute "UPDATE tbl_People SET sysStartTime = " & ts
        .Execute "UPDATE tbl__Comments SET sysStartTime = " & ts
    End With
    
    Fs_Init
    
End Sub

Public Property Get Db_isReadOnly() As Boolean
    Db_isReadOnly = True
    If g_App.App_Mode <> RemoteBackend Then Exit Property
    If g_App.User_Identity = vbNullUser Then Exit Property
    Dim strEditable As String: strEditable = Db_ValueGet("is-editable")
    Db_isReadOnly = Not IsTruthy(strEditable)
End Property

Public Property Get Db_Name() As String

    Db_Name = Db_ValueGet("name")
    If Db_Name <> vbNullString Then Exit Property
    
    If g_App.Model_Current <> vbNullString Then Db_Name = GetBaseName(g_App.Model_Current)
    
End Property

Public Sub Db_Read(src As String)
    
    If src = vbNullString Then Exit Sub
    If Not Db_Validate(src) Then Exit Sub
    
    'g_App.App_ValueSet "modelDb", src
    'g_App.App_ValueRemove "modelDb"              ' Forces 'src' into app_kvs_old rather than app_kvs
    
    Dim sql As String: sql = "INSERT INTO app_Kvs_old (k,SysStartTime,v) VALUES ({k},{t},{v})"
    sql = Replace(sql, "{k}", "'" & "modelDb" & "'")
    sql = Replace(sql, "{v}", "'" & src & "'")
    sql = Replace(sql, "{t}", "#" & SqlDateTime & "#")
    dbLocal.Execute sql
    
End Sub

Public Function Db_ValueGet(s_Key As String) As String
    Db_ValueGet = Nz(ELookup("v", kvsName, "k ='" & s_Key & "'"), vbNullString)
End Function

Public Sub Db_ValueRemove(s_Key As String)
    If Db_isReadOnly Then Exit Sub
    Dim ts As Date
    Dim sql As String: sql = "INSERT INTO " & kvsName & sfx_old & " SELECT *, '{0}' AS SysEndTime FROM " & kvsName & " WHERE k='{1}'"
    Dim rst As Recordset: Set rst = dbLocal.OpenRecordset(kvsName, dbOpenDynaset)
    With rst
        .FindFirst "k='" & s_Key & "'"
        ts = Now
        If Not .NoMatch Then
            sql = Replace(sql, "{0}", SqlDateTime(ts))
            sql = Replace(sql, "{1}", rst!k)
            dbLocal.Execute sql                                         ' dbFailOnError not needed, temporal granularity (e.g. denies > 1 append per minute)
            .Delete
        End If
    End With
    rst.Close
    Set rst = Nothing
End Sub

Public Sub Db_ValueSet(s_Key As String, s_Value As String, Optional forceSet As Boolean)
    If s_Value = vbNullString Then Exit Sub
    If Not forceSet And Db_isReadOnly Then Exit Sub
    Dim ts As Date
    Dim sql As String: sql = "INSERT INTO " & kvsName & sfx_old & " SELECT *, '{0}' AS SysEndTime FROM " & kvsName & " WHERE k='{1}'"
    Dim rst As Recordset: Set rst = dbLocal.OpenRecordset(kvsName, dbOpenDynaset)
    With rst
        .FindFirst "k='" & s_Key & "'"
        ts = Now
        If .NoMatch Then
            .AddNew
        Else
            sql = Replace(sql, "{0}", SqlDateTime(ts))
            sql = Replace(sql, "{1}", rst!k)
            dbLocal.Execute sql                                         ' dbFailOnError not needed, temporal granularity (e.g. denies > 1 append per minute)
            .Edit
        End If
        !k = s_Key
        !v = s_Value
        !sysStartTime = ts
        .Update
        .Bookmark = .LastModified
    End With
    rst.Close
    Set rst = Nothing
    'Debug.Print "Set: '" & s_Key & "' to '" & s_Value & "'"
End Sub

Public Function Db_Validate(src As String) As Boolean
On Error GoTo ValidateDb_Error
    
    Dim db As Database: Set db = Application.DBEngine.OpenDatabase(src, , True)
    
    Dim s_Tbl As String
    Dim tbl As Variant
    For Each tbl In Split(standardTables, ",")
        s_Tbl = CStr(tbl)
        If Not Table_Exists(db, s_Tbl) Then GoTo ExitFunction
    Next tbl
    For Each tbl In Split(temporalTables, ",")
        s_Tbl = CStr(tbl)
        If Not Table_Exists(db, s_Tbl) Then GoTo ExitFunction
        If Not Table_Exists(db, s_Tbl & sfx_old) Then GoTo ExitFunction
    Next tbl
    
    Db_Validate = True

ExitFunction:
    If Not Db_Validate Then MsgBox "The selected file is missing table '[" & s_Tbl & "]'", vbExclamation, "Invalid Model"
    db.Close
    Set db = Nothing
    
On Error GoTo 0
    Exit Function
ValidateDb_Error:
    MsgBox "There was a problem connecting to the selected model", vbExclamation, "Error"
    Set db = Nothing
End Function

Public Function Definition_Archive(field As String, code As String) As Boolean
    
    Definition_Archive = False

On Error GoTo abort
    dbLocal.Execute "INSERT INTO tbl_Definitions_old SELECT * FROM tbl_Definitions WHERE field = '" & field & "' AND code = '" & code & "'", dbFailOnError
    
On Error GoTo rollback
    dbLocal.Execute "UPDATE tbl_Definitions_old SET sysEndTime=#" & SqlDateTime() & "#, sysEndPerson=" & User_Current & " WHERE sysEndTime=#" & vbMaxDate & "# AND field = '" & field & "' AND code = '" & code & "'", dbFailOnError
    Definition_Archive = True
    
abort:
    Exit Function
    
rollback:
    dbLocal.Execute "DELETE * FROM tbl_Definitions_old WHERE sysEndTime=#" & vbMaxDate & "# AND field = '" & field & "' AND code = '" & code & "'"
    MsgBox "The requested change may only be made once per minute. Please wait until the current minute has passed before trying again.", vbExclamation, "Update Denied"
    
End Function

Public Function Definition_Refresh() As Boolean
    Definition_Refresh = False
    Dim fe_CodeCount As Long: fe_CodeCount = DCount("code", "tbl_Definitions_cache")
    Dim fe_TimeMax As Date: fe_TimeMax = Nz(DMax("sysStartTime", "tbl_Definitions_cache"), vbMinDate)
    Dim be_CodeCount As Long: be_CodeCount = DCount("code", "tbl_Definitions")
    Dim be_TimeMax As Date: be_TimeMax = DMax("sysStartTime", "tbl_Definitions")
    If fe_CodeCount <> be_CodeCount Or fe_TimeMax <> be_TimeMax Then
        dbLocal.Execute "DELETE * FROM tbl_Definitions_cache", dbFailOnError
        dbLocal.Execute "INSERT INTO tbl_Definitions_cache SELECT * FROM tbl_Definitions", dbFailOnError
        Debug.Print Now & ": Updated defs cache"
        Definition_Refresh = True
    End If
End Function

Public Function Definition_Update(field As String, code As String, strUpdates As String, Optional ts As Date = vbMaxDate) As Boolean
On Error GoTo abort

    Definition_Update = False
    
    If ts = vbMaxDate Then ts = Now
        
    Dim strSql As String
    If Nz(strUpdates, vbNullString) = vbNullString Then
        strSql = "UPDATE tbl_Definitions SET sysStartTime=#" & SqlDateTime(ts) & "#, sysStartPerson=" & User_Current & ", sysGuid='" & CreateGuid & "' WHERE field = '" & field & "' AND code = '" & code & "'"
    Else
        strSql = "UPDATE tbl_Definitions SET " & strUpdates & ", sysStartTime=#" & SqlDateTime(ts) & "#, sysStartPerson=" & User_Current & ", sysGuid='" & CreateGuid & "' WHERE field = '" & field & "' AND code = '" & code & "'"
    End If
    
    dbLocal.Execute strSql, dbFailOnError
    Definition_Update = True
    Exit Function
    
abort: Stop
End Function

Public Function Dir_Copy(srcPrefix As String, srcDir As Long, trgPrefix As String, trgDir As Long) As Boolean

    Dir_Copy = False
    
    Dim src As String: src = Dir_path(srcPrefix, srcDir)
    Dim trg As String: trg = Dir_path(trgPrefix, trgDir)
    MkDirTree trg
    
    Dim cursor As String: cursor = dir(src & "*.*")
    Do While cursor <> vbNullString
        File_Copy src & cursor, trg & cursor, False, coMutate
        File_SetReadWrite trg & cursor
        cursor = dir()
    Loop
    
    cursor = dir(src, vbDirectory)
    Do While cursor <> vbNullString
        Select Case True
            Case cursor = ".":
            Case cursor = "..":
            Case cursor = "_ss":
            Case (GetAttr(src & cursor) And vbDirectory) = 0:
            Case Else:
                Folder_Copy src & cursor, trg & cursor
                Folder_SetReadWrite trg & cursor
        End Select
        cursor = dir()
    Loop
    
    Dir_Copy = True

End Function

Public Function Dir_Create(prefix As String, dir As Long) As String
    If dir <= 0 Or Not Fs_isConnected Then Exit Function
    MkDirTree Snapshot_path(prefix, dir)
    Dir_Create = Dir_path(prefix, dir)
End Function

Public Sub Dir_Open(prefix As String, dir As Long)
    If Not Fs_isConnected Then
        MsgBox "You must be connected to the filesystem repository in order to open directories", vbInformation, "FsRoot: Not Connected"
        Exit Sub
    End If
    Dim strPath As String: strPath = Dir_Create(prefix, dir)
    Execute strPath
    Comment_Create prefix & CStr(dir), "_clickDir"
End Sub

Public Function Email_Create(mailID As Long) As Long

    Email_Create = 0
    If Db_isReadOnly Then Exit Function
    
    Dim lng_User As Long: lng_User = User_Current
        
    ' Mailstage data - emails to commit
    Dim mailstage As Recordset: Set mailstage = dbLocal.OpenRecordset("SELECT * FROM app_Mailstage WHERE ID = " & mailID & ";", dbOpenSnapshot)
    If mailstage.EOF And mailstage.BOF Then Exit Function
    mailstage.MoveFirst
    
    ' emails data - new E record
    Dim emails As Recordset: Set emails = dbLocal.OpenRecordset("tbl_Emails", dbOpenDynaset)
        emails.AddNew
        emails!subject = mailstage!subject
        emails!from = mailstage!from
        emails!To = mailstage!To
        emails!cc = mailstage!cc
        emails!bcc = mailstage!bcc
        emails!effectiveDate = mailstage!effectiveDate
        emails!effectiveOrg = mailstage!effectiveOrg
        emails!attachments = mailstage!attachments
        emails!hash = mailstage!hash
        emails!hashq = mailstage!hashq
        emails!aliases = mailstage!aliases
        emails!TypeCode = mailstage!TypeCode
        emails!secCode = mailstage!secCode
        emails!classCode = mailstage!classCode
        emails!statusCode = mailstage!statusCode
        emails!sysStartPerson = lng_User
        emails!object = "E" & emails!dir
        emails.Update
        emails.Bookmark = emails.LastModified
        Dir_Create "E", emails!dir
        Email_Create = emails!dir
        
        Dim EmailPath As String: EmailPath = Dir_path("E", emails!dir)
        File_Copy mailstage!fdir & mailstage!fname & "." & mailstage!ftype, EmailPath & emails!object & "." & mailstage!ftype, True, coMutate
        
        Dim bodytext As String: bodytext = mailstage!subject & vbCrLf & vbCrLf & mailstage!bodytext
        File_WriteText bodytext, EmailPath & emails!object & ".txt"
    
    emails.Close
    mailstage.Close
    Set emails = Nothing
    Set mailstage = Nothing
    
End Function

Public Function Email_Matches(hash As String, hashq As String) As String

    Email_Matches = vbNullString
    
    Dim d As Dictionary: Set d = New Dictionary
    Dim dir As String

    Dim rst As Recordset: Set rst = dbLocal.OpenRecordset("SELECT DISTINCT dir FROM quni_Emails WHERE hash = '" & hash & "' OR hashq = '" & hashq & "';", dbOpenForwardOnly)
    Do Until rst.EOF
        dir = rst.Fields("dir").value
        If Not d.Exists(dir) Then d.Add dir, dir
        rst.MoveNext
    Loop
    rst.Close
    
    If d.count > 0 Then
        Dim k As Variant
        For Each k In d.Keys()
            Email_Matches = CsvAdd_numSort(Email_Matches, d(k), True)
        Next
    End If
    
    Set rst = Nothing
    Set d = Nothing
    
End Function

Public Function Email_Stage(strPath As String) As Boolean
On Error GoTo StageEmail_Error

    Email_Stage = False
    
    Dim rst As Recordset: Set rst = dbLocal.OpenRecordset("app_Mailstage", dbOpenDynaset)
    
    Dim FSO As New FileSystemObject
    Dim F As File: Set F = FSO.GetFile(strPath)
    Dim fdir As String: fdir = GetFolderPath(strPath)
    If Len(fdir) > 255 Then GoTo StageEmail_Error
    Dim fname As String: fname = GetBaseName(strPath)
    Dim ftype As String: ftype = GetExtensionName(strPath)
    
    Dim json As String: json = g_App.MailParser_GetJson(strPath)
    Dim d As Dictionary: Set d = ParseJson(json)
    
    Dim hash As String: hash = vbNullString
    If Not IsNull(d("MessageId")) Then
        hash = g_App.Hasher_FromString(d("MessageId"))
    Else
        hash = g_App.Hasher_FromFile(strPath)
    End If
    
    Dim hashq As String: hashq = g_App.Hasher_FromString(d("From") & "_" & d("Date"))
    
    Dim matches As String: matches = Email_Matches(hash, hashq)
    Dim mcount As Long: mcount = 0
    If Not matches = vbNullString Then mcount = ArrayLen(Split(matches, ","))
    
    Dim domain As String: domain = Split(d("From"), "@")(1)
    Dim effectiveOrg As Long: effectiveOrg = Nz(ELookup("dir", "quni_Organisations", "email ALike '%" & domain & "'"), 0)
    
    With rst
        .AddNew
        !fdir = fdir
        !fname = fname
        !ftype = ftype
        If hash <> vbNullString Then !hash = hash
        !hashq = hashq
        !matches = matches
        !mcount = mcount
        !subject = d("Subject")
        !from = d("From")
        !To = d("To")
        !cc = d("Cc")
        !bcc = d("Bcc")
        !effectiveDate = ISODATE(d("Date"))
        If effectiveOrg > 0 Then !effectiveOrg = effectiveOrg
        !attachments = d("Attachments")
        !bodytext = d("BodyText")
        .Update
        .Bookmark = .LastModified
    End With
    
    Email_Stage = True
 
StageEmail_Error:
    rst.Close
    Set rst = Nothing
    Set FSO = Nothing
    Set F = Nothing
    Set d = Nothing
End Function

Public Function Entity_Archive(tblName As String, dir As Long) As Boolean

    Entity_Archive = False
    
On Error GoTo abort
    dbLocal.Execute "INSERT INTO " & tblName & sfx_old & " SELECT * FROM " & tblName & " WHERE dir=" & dir, dbFailOnError

On Error GoTo rollback
    dbLocal.Execute "UPDATE " & tblName & sfx_old & " SET sysEndTime=#" & SqlDateTime() & "#, sysEndPerson=" & User_Current & " WHERE sysEndTime=#" & vbMaxDate & "# AND dir=" & dir, dbFailOnError
    Entity_Archive = True

abort:
    Exit Function

rollback:
    dbLocal.Execute "DELETE * FROM " & tblName & sfx_old & " WHERE sysEndTime=#" & vbMaxDate & "# AND dir=" & dir
    MsgBox "The requested change may only be made once per minute. Please wait until the current minute has passed before trying again.", vbExclamation, "Update Denied"
    
End Function

Public Function Entity_ArchiveAndUpdate(tblName As String, dir As Long, strUpdates As String, Optional ts As Date = vbMaxDate) As Boolean

    Entity_ArchiveAndUpdate = False
        
On Error GoTo abort
    dbLocal.Execute "INSERT INTO " & tblName & sfx_old & " SELECT * FROM " & tblName & " WHERE dir=" & dir, dbFailOnError
    
On Error GoTo rollback
    If ts = vbMaxDate Then ts = Now
    dbLocal.Execute "UPDATE " & tblName & sfx_old & " SET sysEndTime=#" & SqlDateTime(ts) & "#, sysEndPerson=" & User_Current & " WHERE sysEndTime=#" & vbMaxDate & "# AND dir=" & dir, dbFailOnError
        
    Dim strSql As String
    If Nz(strUpdates, vbNullString) = vbNullString Then
        strSql = "UPDATE " & tblName & " SET sysStartTime=#" & SqlDateTime(ts) & "#, sysStartPerson=" & User_Current & ", sysGuid='" & CreateGuid & "' WHERE dir=" & dir
    Else
        strSql = "UPDATE " & tblName & " SET " & strUpdates & ", sysStartTime=#" & SqlDateTime(ts) & "#, sysStartPerson=" & User_Current & ", sysGuid='" & CreateGuid & "' WHERE dir=" & dir
    End If
    
    dbLocal.Execute strSql, dbFailOnError
    Entity_ArchiveAndUpdate = True

abort:
    Exit Function
    
rollback:
    dbLocal.Execute "DELETE * FROM " & tblName & sfx_old & " WHERE sysEndTime=#" & vbMaxDate & "# AND dir=" & dir
    MsgBox "The requested change may only be made once per minute. If it is important, please wait until the current minute has passed and then try again.", vbExclamation, "Update Denied"
    
End Function

Public Sub Entity_Undelete(tblName As String, dir As Long)
    dbLocal.Execute "INSERT INTO " & tblName & " SELECT TOP 1 * FROM " & tblName & sfx_old & " WHERE dir=" & dir & " ORDER BY sysEndTime DESC;", dbFailOnError
    dbLocal.Execute "UPDATE " & tblName & " SET sysStartTime=#" & SqlDateTime & "#, sysStartPerson=" & User_Current & ", sysEndTime=#" & vbMaxDate & "#, sysEndPerson=0 WHERE dir=" & dir, dbFailOnError
End Sub

Public Function Entity_Update(tblName As String, dir As Long, strUpdates As String, Optional ts As Date = vbMaxDate) As Boolean
On Error GoTo abort

    Entity_Update = False
    
    If ts = vbMaxDate Then ts = Now
    
    Dim strSql As String
    If Nz(strUpdates, vbNullString) = vbNullString Then
        strSql = "UPDATE " & tblName & " SET sysStartTime=#" & SqlDateTime(ts) & "#, sysStartPerson=" & User_Current & ", sysGuid='" & CreateGuid & "' WHERE dir=" & dir
    Else
        strSql = "UPDATE " & tblName & " SET " & strUpdates & ", sysStartTime=#" & SqlDateTime(ts) & "#, sysStartPerson=" & User_Current & ", sysGuid='" & CreateGuid & "' WHERE dir=" & dir
    End If
    
    dbLocal.Execute strSql, dbFailOnError
    Entity_Update = True
    Exit Function
    
abort: Stop
End Function

Public Function File_Matches(hash As String, hashq As String, Optional refNo As String = vbNullString) As String

    File_Matches = vbNullString
    
    Dim d As Dictionary: Set d = New Dictionary
    Dim dir As String

    Dim rst As Recordset: Set rst = dbLocal.OpenRecordset("SELECT DISTINCT target FROM tbl__Files WHERE hash = '" & hash & "' OR hashq = '" & hashq & "';", dbOpenForwardOnly)
    Do Until rst.EOF
        dir = rst.Fields("target").value
        If Not d.Exists(dir) Then d.Add dir, dir
        rst.MoveNext
    Loop
    rst.Close
    
    If Not refNo = vbNullString Then
        Set rst = dbLocal.OpenRecordset("SELECT dir FROM quni_Inputs WHERE refNo = '" & refNo & "';", dbOpenForwardOnly)
        Do Until rst.EOF
            dir = rst.Fields("dir").value
            If Not d.Exists(dir) Then d.Add dir, dir
            rst.MoveNext
        Loop
        rst.Close
    End If
    
    If d.count > 0 Then
        Dim k As Variant
        For Each k In d.Keys()
            File_Matches = CsvAdd_numSort(File_Matches, d(k), True)
        Next
    End If
    
    Set rst = Nothing
    Set d = Nothing
    
End Function

Public Function File_Stage(strPath As String) As Boolean
On Error GoTo StageFile_Error

    File_Stage = False
    
    Dim rst As Recordset: Set rst = dbLocal.OpenRecordset("app_Filestage", dbOpenDynaset)
    
    Dim FSO As New FileSystemObject
    Dim F As File: Set F = FSO.GetFile(strPath)
    Dim fdir As String: fdir = GetFolderPath(strPath)
    If Len(fdir) > 255 Then GoTo StageFile_Error
        
    Dim fname As String: fname = GetBaseName(strPath)
    Dim ftype As String: ftype = GetExtensionName(strPath)
    
    Dim dmin As Date: dmin = mindate(F.DateLastModified, F.DateCreated)
    Dim hashq As String: hashq = g_App.Hasher_FromString(Date2Long(dmin) & ftype & F.Size)
    Dim hash As String: hash = g_App.Hasher_FromFile(strPath)
    Dim matches As String: matches = File_Matches(hash, hashq)
    
    Dim mcount As Long: mcount = 0
    If Not matches = vbNullString Then mcount = ArrayLen(Split(matches, ","))

    With rst
        .AddNew
        !fdir = fdir
        !fname = fname
        !ftype = ftype
        !fsize = F.Size
        !fDateCreated = F.DateCreated
        !fDateLastModified = F.DateLastModified
        !hashq = hashq
        !hash = hash
        !matches = matches
        !mcount = mcount
        !Title = fname
        .Update
        .Bookmark = .LastModified
    End With
    
    File_Stage = True
    
StageFile_Error:
    rst.Close
    Set rst = Nothing
    Set FSO = Nothing
    Set F = Nothing
End Function

Public Function Fs_Init() As Boolean

    Dim strId As String: strId = Db_ID
    Dim strName As String: strName = Db_Name
    
    If strId = vbNullString Or strName = vbNullString Then GoTo abort
    
    Dim d As Dictionary: Set d = New Dictionary
    d.Add "id", Db_ID
    d.Add "model", Db_Name
    Dim json As String: json = ConvertToJson(d, 4)
        
    Dim strPath As String: strPath = g_App.App_Models & Db_Name & vbBackSlash
    If Not Folder_Exists(strPath) Then MkDirTree strPath
    If Folder_Exists(strPath) Then File_WriteText json, strPath & ".fmms"
    If Not File_Exists(strPath & ".fmms") Then GoTo abort
        
    Exit Function
    
abort:
    MsgBox "Could not create new Fs repository.", vbInformation, "Create FsRoot"
    Stop
End Function

Public Property Get Fs_isConnected() As Boolean
    Fs_isConnected = g_App.App_ValueGet("modelFs") <> vbNullString
End Property

Public Property Get Fs_Root() As String
    Fs_Root = g_App.App_ValueGet("modelFs")
End Property

Public Sub Fs_Refresh()
    Dim newPath As String: newPath = g_App.App_Models & Db_Name & vbBackSlash
    If Fs_Validate(newPath) Then GoTo Valid

    newPath = CurrentProject.Path & vbBackSlash & Db_Name & vbBackSlash
    If Fs_Validate(newPath) Then GoTo Valid
    
    g_App.App_ValueRemove "modelFs"
    Exit Sub
    
Valid:
    g_App.App_ValueSet "modelFs", newPath
End Sub

Public Function Fs_Validate(Optional strPath As String = vbNullString) As Boolean
On Error Resume Next
        
    If strPath = vbNullString Then strPath = Fs_Root

    Dim strJson As String: strJson = File_ReadText(strPath & ".fmms")
    If strJson = vbNullString Then
        Debug.Print "Error: Could not locate this model's repository"
        Exit Function
    End If

    Dim objJson As Object: Set objJson = mod_JsonUtils.ParseJson(strJson)
    If objJson Is Nothing Then
        Debug.Print "Error: The specified FMMS repo information is corrupt or malformed"
        Exit Function
    End If

    Dim strId As String: strId = Nz(objJson("id"), vbNullString)
    If strId = vbNullString Then
        Debug.Print "Error: The specified FMMS repo information doesn't contain a valid model id"
        Exit Function
    End If

    Fs_Validate = strId = Db_ID
    If Not Fs_Validate Then Debug.Print "Error: The FMMS repo is not the right one for this model"

End Function

Public Function Hasher_Validate() As Boolean
    Dim hash As String: hash = g_App.Hasher_FromString(vbNullString)         ' Null key for validation (could be anything)
    Dim mv_hash As String: mv_hash = Db_ValueGet("valid-hash")
    If mv_hash = vbNullString Then
        Db_ValueSet "valid-hash", hash
        mv_hash = Db_ValueGet("valid-hash")
    End If
    Hasher_Validate = hash = mv_hash
End Function

Public Function Input_Commit(fileList As String) As Long

    Input_Commit = 0
    If Db_isReadOnly Then Exit Function
    
    Dim lng_User As Long: lng_User = User_Current
        
    ' Filestage data - files to commit
    Dim filestage As Recordset: Set filestage = dbLocal.OpenRecordset("SELECT * FROM app_Filestage WHERE ID IN(" & fileList & ");", dbOpenSnapshot)
    If filestage.EOF And filestage.BOF Then Exit Function
    filestage.MoveFirst
    
    ' Inputs data - new N record
    Dim inputs As Recordset: Set inputs = dbLocal.OpenRecordset("tbl_Inputs", dbOpenDynaset)
        inputs.AddNew
        inputs!Title = filestage!Title
        inputs!aliases = filestage!aliases
        inputs!TypeCode = filestage!TypeCode
        inputs!secCode = filestage!secCode
        inputs!classCode = filestage!classCode
        inputs!statusCode = filestage!statusCode
        inputs!refNo = filestage!refNo
        inputs!revNo = filestage!revNo
        inputs!effectiveDate = filestage!effectiveDate
        inputs!effectiveOrg = filestage!effectiveOrg
        inputs!sysStartPerson = lng_User
        inputs!object = "N" & inputs!dir
        inputs.Update
        inputs.Bookmark = inputs.LastModified
        Dir_Create "N", inputs!dir
        Input_Commit = inputs!dir
        
        Dim inputPath As String: inputPath = Dir_path("N", inputs!dir)
        
    ' File data - historic record
    Dim files As Recordset: Set files = dbLocal.OpenRecordset("tbl__Files", dbOpenDynaset)
    Do Until filestage.EOF
        files.AddNew
        files!fdir = filestage!fdir
        files!fname = filestage!fname
        files!ftype = filestage!ftype
        files!fsize = filestage!fsize
        files!fDateCreated = filestage!fDateCreated
        files!fDateLastModified = filestage!fDateLastModified
        files!hash = filestage!hash
        files!hashq = filestage!hashq
        files!target = inputs!dir
        files!sysStartPerson = lng_User
        files.Update
        files.Bookmark = files.LastModified
        File_Copy filestage!fdir & filestage!fname & "." & filestage!ftype, inputPath & filestage!fname & "." & filestage!ftype, True, coMutate
        filestage.MoveNext
    Loop
    
    files.Close
    inputs.Close
    filestage.Close
    Set files = Nothing
    Set inputs = Nothing
    Set filestage = Nothing
    
End Function

Public Function Input_Merge(grouping As Long, fileList As String) As Long

    Input_Merge = 0
    If Db_isReadOnly Then Exit Function
    
    Dim lng_User As Long: lng_User = User_Current
        
    ' Filestage - get files to commit
    Dim filestage As Recordset: Set filestage = dbLocal.OpenRecordset("SELECT * FROM app_Filestage WHERE ID IN(" & fileList & ");", dbOpenSnapshot)
    If filestage.EOF And filestage.BOF Then Exit Function
    filestage.MoveFirst
        
    ' Files - snapshot existing then merge new files
    Snapshot_Create "N", grouping, "Auto snapshot prior to merge from filestage.", True
    Dim inputPath As String: inputPath = Dir_path("N", grouping)
    
    Dim files As Recordset: Set files = dbLocal.OpenRecordset("tbl__Files", dbOpenDynaset)
    Do Until filestage.EOF
        File_Copy filestage!fdir & filestage!fname & "." & filestage!ftype, inputPath & filestage!fname & "." & filestage!ftype, True, coMutate
        files.AddNew
        files!fdir = filestage!fdir
        files!fname = filestage!fname
        files!ftype = filestage!ftype
        files!fsize = filestage!fsize
        files!fDateCreated = filestage!fDateCreated
        files!fDateLastModified = filestage!fDateLastModified
        files!hash = filestage!hash
        files!hashq = filestage!hashq
        files!target = grouping
        files!sysStartPerson = lng_User
        files.Update
        files.Bookmark = files.LastModified
        filestage.MoveNext
    Loop
    
    ' Inputs data - archive then update existing N record
    Entity_Archive "tbl_Inputs", grouping
    
    filestage.MoveFirst
    Dim inputs As Recordset: Set inputs = dbLocal.OpenRecordset("SELECT * FROM tbl_Inputs WHERE dir=" & grouping & ";", dbOpenDynaset)
        inputs.Edit
        inputs!Title = filestage!Title
        inputs!aliases = filestage!aliases
        inputs!TypeCode = filestage!TypeCode
        inputs!secCode = filestage!secCode
        inputs!classCode = filestage!classCode
        inputs!statusCode = filestage!statusCode
        inputs!refNo = filestage!refNo
        inputs!revNo = filestage!revNo
        inputs!effectiveDate = filestage!effectiveDate
        inputs!effectiveOrg = filestage!effectiveOrg
        inputs!sysStartPerson = lng_User
        'inputs!sysStartTime = Now
        inputs.Update
        inputs.Bookmark = inputs.LastModified
        
        Input_Merge = grouping
        
    files.Close
    inputs.Close
    filestage.Close
    Set files = Nothing
    Set inputs = Nothing
    Set filestage = Nothing
    
End Function

Public Function MailParser_Validate() As Boolean
    MailParser_Validate = File_Exists(g_App.App_Assets & "mail-parser.exe")
End Function

Public Function Memorandum_Duplicate(srcId As Long) As Long

    Memorandum_Duplicate = 0
    If Db_isReadOnly Then Exit Function
    
On Error GoTo CreateMemorandum_Error
    
    Dim src As Recordset: Set src = dbLocal.OpenRecordset("tbl_Memoranda", dbOpenSnapshot)
    src.FindFirst "dir=" & srcId
    If Not src.NoMatch Then
        Dim rst As Recordset: Set rst = dbLocal.OpenRecordset("tbl_Memoranda", dbOpenDynaset)
            rst.AddNew
            rst!link = src!link
            rst!sysStartPerson = User_Current
            rst!Title = src!Title
            rst!progress = 0
            rst!origCode = src!origCode
            rst!sysCode = src!sysCode
            rst!locCode = src!locCode
            rst!TypeCode = src!TypeCode
            rst!roleCode = src!roleCode
            rst!secCode = src!secCode
            rst!classCode = src!classCode
            rst!location = src!location
            rst!object = "M" & rst!dir
            rst.Update
            rst.Bookmark = rst.LastModified
            Dir_Copy "M", src!dir, "M", rst!dir
            
            Memorandum_Duplicate = rst!dir
    End If
    
On Error GoTo 0
    src.Close
    rst.Close
    Set src = Nothing
    Set rst = Nothing
    Exit Function
    
CreateMemorandum_Error:
    
End Function

Public Function Organisation_Create(strName As String) As Long
    
    Organisation_Create = 0
    If Db_isReadOnly Then Exit Function
    
On Error GoTo CreateOrganisation_Error

    Dim rst As Recordset: Set rst = dbLocal.OpenRecordset("tbl_Organisations", dbOpenDynaset)
        rst.AddNew
        rst!Title = strName
        rst!sysStartPerson = User_Current
        rst!object = "O" & rst!dir
        rst.Update
        rst.Bookmark = rst.LastModified
        
        Dir_Create "O", rst!dir
        
    Organisation_Create = rst!dir
    
On Error GoTo 0
    rst.Close
    Set rst = Nothing
    Exit Function
    
CreateOrganisation_Error:

End Function

Public Function Output_Duplicate(srcId As Long) As Long
    
    Output_Duplicate = 0
    If Db_isReadOnly Then Exit Function
    
On Error GoTo CreateOutput_Error
    
    Dim src As Recordset: Set src = dbLocal.OpenRecordset("tbl_Outputs", dbOpenSnapshot)
    src.FindFirst "dir=" & srcId
    If Not src.NoMatch Then
        Dim rst As Recordset: Set rst = dbLocal.OpenRecordset("tbl_Outputs", dbOpenDynaset)
            rst.AddNew
            rst!link = src!link
            rst!sysStartPerson = User_Current
            rst!Title = src!Title
            rst!scheme = Form_frm_Schemes.CurrentScheme
            rst!progress = 0
            rst!origCode = src!origCode
            rst!sysCode = src!sysCode
            rst!locCode = src!locCode
            rst!TypeCode = src!TypeCode
            rst!roleCode = src!roleCode
            rst!revCode = src!revCode
            rst!secCode = src!secCode
            rst!classCode = src!classCode
            rst!object = "U" & rst!dir
            'rst!sysStartTime = src!
            'rst!isFrozen = src!
            'rst!aliases = src!
            'rst!attachedTo = src!
            'rst!refNo = src!
            'rst!revNo = src!
            'rst!effectiveDate = src!
            'rst!statusCode = src!statusCode
            'rst!wsCode = src!wsCode
            'rst!responsible = src!
            'rst!sysEndTime = src!
            'rst!sysEndPerson = src!
            rst.Update
            rst.Bookmark = rst.LastModified
            Dir_Copy "U", src!dir, "U", rst!dir
            
            Output_Duplicate = rst!dir
    End If
    
On Error GoTo 0
    src.Close
    rst.Close
    Set src = Nothing
    Set rst = Nothing
    Exit Function
    
CreateOutput_Error:
    
End Function

Public Function Person_Create(strName As String) As Long
    
    Person_Create = 0
    If Db_isReadOnly Then Exit Function
    
On Error GoTo CreatePerson_Error

    Dim rst As Recordset: Set rst = dbLocal.OpenRecordset("tbl_People", dbOpenDynaset)
        rst.AddNew
        rst!Title = strName
        rst!sysStartPerson = User_Current
        rst!object = "P" & rst!dir
        rst.Update
        rst.Bookmark = rst.LastModified
        
        Dir_Create "P", rst!dir
        
    Person_Create = rst!dir
    
On Error GoTo 0
    rst.Close
    Set rst = Nothing
    Exit Function
    
CreatePerson_Error:

End Function

Public Function Scheme_Children(dir As Long) As String
    If Nz(dir, 0) <= 0 Then GoTo error
    Dim scheme As String: scheme = Nz(ELookup("schemeCode", "tbl_Schemes", "dir=" & dir), 0)
    If Nz(scheme, 0) <= 0 Then GoTo error
    Dim csv As String: csv = vbNullString
    Dim rst As Recordset: Set rst = dbLocal.OpenRecordset("SELECT dir, sort, indent FROM tbl_Schemes WHERE schemeCode='" & scheme & "' ORDER BY sort;", dbOpenSnapshot)
    Dim indent As Integer
    rst.FindFirst "dir=" & dir
    If rst!indent > 1 Then GoTo Done
    indent = rst!indent
    rst.MoveNext
    Do Until rst.EOF
        If rst!indent <= indent Then GoTo Done
        csv = csv & rst!dir & ","
        rst.MoveNext
    Loop
Done:
    rst.Close
    Set rst = Nothing
    Scheme_Children = mod_StringUtils.TrimTrailingChr(csv, ",")
    If Scheme_Children <> vbNullString Then Exit Function
error:
    Scheme_Children = "0"
End Function

Public Function Scheme_Parents(dir As Long) As String
    If Nz(dir, 0) <= 0 Then GoTo error
    Dim scheme As String: scheme = Nz(ELookup("schemeCode", "tbl_Schemes", "dir=" & dir), 0)
    If Nz(scheme, 0) <= 0 Then GoTo error
    Dim csv As String: csv = vbNullString
    Dim rst As Recordset: Set rst = dbLocal.OpenRecordset("SELECT dir, sort, indent FROM tbl_Schemes WHERE schemeCode='" & scheme & "' ORDER BY sort DESC;", dbOpenSnapshot)
    rst.FindFirst "dir=" & dir
    Dim indent As Integer: indent = rst!indent
    Do Until rst.EOF
        If indent = 0 Then GoTo Done
        If rst!indent <> indent Then
            csv = csv & rst!dir & ","
            indent = rst!indent
        End If
        rst.FindNext "indent=" & rst!indent - 1
    Loop
Done:
    rst.Close
    Set rst = Nothing
    Scheme_Parents = mod_StringUtils.TrimTrailingChr(csv, ",")
    If Scheme_Parents <> vbNullString Then Exit Function
error:
    Scheme_Parents = "0"
End Function

'Public Function Scheme_Parent(dir As Long) As Long
'    If Nz(dir, 0) <= 0 Then GoTo error
'    Dim scheme As String: scheme = Nz(ELookup("schemeCode", "tbl_Schemes", "dir=" & dir), 0)
'    If Nz(scheme, 0) <= 0 Then GoTo error
'    Dim rst As Recordset: Set rst = dbLocal.OpenRecordset("SELECT dir, sort, indent FROM tbl_Schemes WHERE schemeCode='" & scheme & "' ORDER BY sort DESC;", dbOpenSnapshot)
'    rst.FindFirst "dir=" & dir
'    Dim indent As Integer: indent = rst!indent
'    Do Until rst.EOF
'        If indent = 0 Then GoTo error
'        If rst!indent <> indent Then GoTo Done
'        rst.FindNext "indent=" & rst!indent - 1
'    Loop
'Done:
'    Scheme_Parent = Nz(rst!dir, 0)
'    rst.Close
'    Set rst = Nothing
'    Exit Function
'error:
'    Scheme_Parent = 0
'End Function

Public Sub Scheme_RefreshWbs(schemeCode As String)
    Dim rst As Recordset: Set rst = dbLocal.OpenRecordset("SELECT sort, indent, wbs FROM tbl_Schemes WHERE schemeCode='" & schemeCode & "' ORDER BY sort;", dbOpenDynaset)
    rst.MoveLast
    rst.MoveFirst
    Dim o(3), n(3) As Long
    Dim wbs As String
    Dim sort As Long: sort = 1
    Do While Not rst.EOF
        Select Case Nz(rst!indent, 0)
            Case 2:
                n(0) = o(0)
                n(1) = o(1)
                n(2) = o(2) + 1
            Case 1:
                n(0) = o(0)
                n(1) = o(1) + 1
                n(2) = 0
            Case Else:
                n(0) = o(0) + 1
                n(1) = 0
                n(2) = 0
        End Select
        rst.Edit
            rst!wbs = vbNullString
            If n(0) > 0 Then rst!wbs = schemeCode + "-" + CStr(n(0))
            If n(1) > 0 Then rst!wbs = rst!wbs + "." + CStr(n(1))
            If n(2) > 0 Then rst!wbs = rst!wbs + "." + CStr(n(2))
            rst!sort = sort
        rst.Update
        o(0) = n(0)
        o(1) = n(1)
        o(2) = n(2)
        rst.MoveNext
        sort = sort + 1
    Loop
    rst.Close
    Set rst = Nothing
End Sub

Public Sub Snapshot_Browse(prefix As String, dirId As Long)
    If dirId <= 0 Or Not Fs_isConnected Then Exit Sub
    Dim fpath As String: fpath = Snapshot_path(prefix, dirId)
    If Not Folder_Exists(fpath) Then MkDirTree fpath
    Execute fpath
End Sub

Public Sub Snapshot_Create(prefix As String, dirId As Long, Optional msg As String = vbNullString, Optional skipPrompts As Boolean = False, Optional location As String, Optional sysGuid As String)
    
    If dirId <= 0 Or Not Fs_isConnected Then Exit Sub
    
    Dim ts As Date: ts = Now
    Dim src As String: src = Dir_path(prefix, dirId)
    Dim trg As String: trg = Snapshot_path(prefix, dirId, ToYYYYMMDDHHNN(ts))
    
    ' Create Snapshot
    MkDirTree trg
    Dim cursor As String: cursor = dir(src & "*.*")
    Do While cursor <> vbNullString
        File_Copy src & cursor, trg & cursor, False, coSkip
        cursor = dir()
    Loop
    cursor = dir(src, vbDirectory)
    Do While cursor <> vbNullString
        Select Case True
            Case cursor = ".":
            Case cursor = "..":
            Case cursor = "_ss":
            Case (GetAttr(src & cursor) And vbDirectory) = 0:
            Case Else: Folder_Copy src & cursor, trg & cursor
        End Select
        cursor = dir()
    Loop
    
    ' Create ssInfo file
    Dim userId As Long: userId = User_Current
    If Nz(msg, vbNullString) = vbNullString Then msg = "Not entered"
    If Nz(location, vbNullString) = vbNullString Then location = "N/A"
    If Nz(sysGuid, vbNullString) = vbNullString Then sysGuid = "N/A"
    
    Dim fname As String: fname = trg & ssInfo
    If File_Exists(fname) Then File_Delete fname, True
    
    Dim d As Dictionary: Set d = New Dictionary
    d.Add "object", prefix & dirId
    d.Add "date", SqlDateTime(ts)
    d.Add "user", "P" & userId & " / " & Nz(ELookup("title", "quni_People", "dir=" & userId), "Unknown")
    d.Add "note", msg
    d.Add "location", location
    d.Add "sysGuid", sysGuid
    File_WriteText ConvertToJson(d, 4), fname
    Set d = Nothing
    
    Folder_SetReadOnly trg
    
    If skipPrompts Then Exit Sub
    If MsgBox("Click OK to view the snapshot now?", vbOKCancel + vbQuestion, prefix & dirId & " - Snapshot Created") = vbOK Then Execute trg
    
End Sub

Public Sub Snapshot_Open(prefix As String, dirId As Long, Optional ts As Date = vbMinDate)
    If dirId <= 0 Or Not Fs_isConnected Then Exit Sub
    If ts = vbMinDate Then ts = Now
    Dim ssDir As String: ssDir = Snapshot_path(prefix, dirId)
    Dim cursor As String: cursor = dir(ssDir, vbDirectory)
    Dim alist As Object: Set alist = CreateObject("System.Collections.ArrayList")
    Do While cursor <> vbNullString
        Select Case True
            Case cursor = ".":
            Case cursor = "..":
            Case (GetAttr(ssDir & cursor) And vbDirectory) = 0:
            Case Else: alist.Add cursor
        End Select
        cursor = dir()
    Loop
    alist.sort
    alist.Reverse
    
    Dim dblTs As Double: dblTs = ToYYYYMMDDHHNN(ts)
    
    Dim i As Long
    For i = 0 To alist.count - 1
        If alist(i) - dblTs <= 0 Then GoTo Done
    Next
    
    Dim msg As String: msg = "Could not locate any snapshots prior to " & ts & "."
    If ts = vbMaxDate Then msg = "Could not locate any snapshots for this record."
    If MsgBox(msg & vbCrLf & vbCrLf & "Click OK to browse all snapshots now?", vbInformation + vbOKCancel, prefix & dirId & " Snapshot - Not Found!") = vbOK Then Snapshot_Browse prefix, dirId
    Exit Sub
    
Done:
    Execute ssDir & alist(i)
End Sub

Public Function Url_Stage(strUrl As String) As Boolean
On Error GoTo StageUrl_Error
    
    Url_Stage = False
    
    Dim rst As Recordset: Set rst = dbLocal.OpenRecordset("app_Filestage", dbOpenDynaset)
    
    Dim hashq As String: hashq = g_App.Hasher_FromString(strUrl)
    Dim strPath As String: strPath = g_App.User_TempFiles & hashq & ".url"
    With CreateObject("WScript.Shell").CreateShortcut(strPath)
        .TargetPath = strUrl
        .Save
    End With
    Dim FSO As New FileSystemObject
    Dim F As File: Set F = FSO.GetFile(strPath)
    Dim fdir As String: fdir = GetFolderPath(strPath)
    If Len(fdir) > 255 Then GoTo StageUrl_Error
    
    Dim fname As String: fname = GetBaseName(strPath)
    Dim ftype As String: ftype = GetExtensionName(strPath)
    Dim hash As String: hash = g_App.Hasher_FromFile(strPath)
    Dim matches As String: matches = File_Matches(hash, hashq)
    
    Dim mcount As Long: mcount = 0
    If Not matches = vbNullString Then mcount = ArrayLen(Split(matches, ","))
    
    With rst
        .AddNew
        !fdir = fdir
        !fname = fname
        !ftype = ftype
        !fsize = F.Size
        !fDateCreated = F.DateCreated
        !fDateLastModified = F.DateLastModified
        !hashq = hashq
        !hash = hash
        !matches = matches
        !mcount = mcount
        !Title = strUrl
        .Update
        .Bookmark = .LastModified
    End With
    
    Url_Stage = True
    
StageUrl_Error:
    rst.Close
    Set rst = Nothing
    Set FSO = Nothing
    Set F = Nothing
End Function

Public Property Get User_Current() As Long
    User_Current = Nz(ELookup("dir", "tbl_People", "identity='" & g_App.User_Identity & "'", "dir ASC"), 0)
End Property

Public Property Get User_CurrentSchemes() As String
    User_CurrentSchemes = Nz(ELookup("attachedTo", "tbl_People", "dir=" & User_Current), vbNullString)
End Property

Public Property Get User_IsBlocked() As Boolean
    User_IsBlocked = Nz(ELookup("isBlocked", "quni_People", "dir=" & User_Current), False)
End Property

Public Function User_IsManager() As Boolean
    User_IsManager = False
    If User_Current <= 0 Then Exit Function
    User_IsManager = InStr(User_Permissions, ",*,") > 0
End Function

Public Function WordPad_ArchiveAndUpdate(object As String, strUpdates As String, Optional ts As Date = vbMaxDate) As Boolean
    
    WordPad_ArchiveAndUpdate = False
    
On Error GoTo abort
    dbLocal.Execute "INSERT INTO tbl_WordPad_old SELECT * FROM tbl_WordPad WHERE object = '" & object & "'", dbFailOnError

On Error GoTo rollback
    
    If ts = vbMaxDate Then ts = Now
    dbLocal.Execute "UPDATE tbl_WordPad_old SET sysEndTime=#" & SqlDateTime(ts) & "#, sysEndPerson=" & User_Current & " WHERE sysEndTime=#" & vbMaxDate & "# AND object = '" & object & "'", dbFailOnError
        
    Dim strSql As String
    If Nz(strUpdates, vbNullString) = vbNullString Then
        strSql = "UPDATE tbl_WordPad SET sysStartTime=#" & SqlDateTime(ts) & "#, sysStartPerson=" & User_Current & ", sysGuid='" & CreateGuid & "' WHERE object = '" & object & "'"
    Else
        strSql = "UPDATE tbl_WordPad SET " & strUpdates & ", sysStartTime=#" & SqlDateTime(ts) & "#, sysStartPerson=" & User_Current & ", sysGuid='" & CreateGuid & "' WHERE object = '" & object & "'"
    End If
    
    dbLocal.Execute strSql, dbFailOnError
    WordPad_ArchiveAndUpdate = True
    
abort:
    Exit Function
    
rollback:
    Debug.Print strSql
    dbLocal.Execute "DELETE * FROM tbl_WordPad_old WHERE sysEndTime=#" & vbMaxDate & "# AND object = '" & object & "'"
    MsgBox "The requested change may only be made once per minute. If it is important, please wait until the current minute has passed and then try again.", vbExclamation, "Update Denied"
    
End Function

Public Function WordPad_Create(strObj As String) As Boolean
    
    WordPad_Create = False
    If Db_isReadOnly Then Exit Function
    
On Error GoTo CreateWordPad_Error

    Dim rst As Recordset: Set rst = dbLocal.OpenRecordset("tbl_WordPad", dbOpenDynaset)
        rst.AddNew
        rst!object = strObj
        rst!sysStartPerson = User_Current
        rst!wp = Db_ValueGet("wp-" & Left(strObj, 1))
        rst.Update
        rst.Bookmark = rst.LastModified
                
    WordPad_Create = True
    
On Error GoTo 0
    rst.Close
    Set rst = Nothing
    Exit Function
    
CreateWordPad_Error:

End Function

