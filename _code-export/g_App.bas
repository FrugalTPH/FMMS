Attribute VB_Name = "g_App"
Option Explicit
Private Const kvsName As String = "app_Kvs"
Private auth_AccessToken As String
Private auth_ExpiresOn As Date
Private auth_UserEmail As String
Private auth_UserId As String
Private auth_UserName As String
Private auth__api As String
Private auth__clientCallback As String
Private auth__clientId As String
Private auth__clientSecret As String
Private auth__server As String
Private e_Mode As AppMode
Private s_Model As String


'''''''''''''
'' PRIVATE ''
'''''''''''''


Private Property Get App_InstallRoot() As String
    'App_InstallRoot = Environ$("FMMS") & vbBackSlash
    App_InstallRoot = CurrentProject.Path & vbBackSlash
End Property

Private Sub App_RestoreTable(src As String, structOnly As Boolean)
    If Not Table_Exists(dbLocal, src) Then DoCmd.TransferDatabase acImport, "Microsoft Access", App_Scripts & "default.accdb", acTable, src, src, structOnly
End Sub


''''''''''''
'' PUBLIC ''
''''''''''''


Public Sub Initialize()
    e_Mode = Uninitialized
    App_RestoreTable kvsName, True
    App_RestoreTable kvsName & sfx_old, True
    App_RestoreTable "app_Filestage", True
    App_RestoreTable "app_Mailstage", True
    MkDirTree User_TempFiles
    If App_isLocalSnapshot Then
        e_Mode = LocalSnapshot
    Else
        e_Mode = RemoteBackend
        'Auth User
    End If
End Sub

Public Sub Templates_Open()
    Execute App_Templates
End Sub

Public Sub Terminate()
    dbLocal True
    DoEvents
    'Access.Quit acQuitSaveAll
End Sub

Public Property Get App_Assets() As String
    App_Assets = App_InstallRoot & "assets" & vbBackSlash
End Property

Public Property Get App_isLocalSnapshot() As Boolean
On Error Resume Next
    App_isLocalSnapshot = Len(dbLocal.TableDefs("tbl_Model").Connect) <= 0
End Property

Public Property Get App_Mode() As String
    App_Mode = e_Mode
End Property

Public Property Get App_Models() As String
    App_Models = App_InstallRoot & "models" & vbBackSlash
End Property

Public Sub App_PropertySet(strProperty As String, newValue As String)
On Error GoTo ErrorHandler
    dbLocal.Properties(strProperty) = newValue
    dbLocal.Properties.Refresh
    Exit Sub

ErrorHandler:
    If Err.Number = 3270 Then
        Dim obj As Object: Set obj = dbLocal.CreateProperty(strProperty, dbText, newValue)
        dbLocal.Properties.Append obj
        Set obj = Nothing
        dbLocal.Properties.Refresh
    End If
    Resume Next
End Sub

Public Property Get App_Scripts() As String
    App_Scripts = App_InstallRoot & "scripts" & vbBackSlash
End Property

Public Property Get App_Templates() As String
    App_Templates = App_InstallRoot & "_templates" & vbBackSlash
End Property

Public Function App_ValueGet(s_Key As String) As String
    App_ValueGet = Nz(ELookup("v", kvsName, "k ='" & s_Key & "'"), vbNullString)
End Function

Public Sub App_ValueRemove(s_Key As String)
    Dim ts As Date
    Dim sql As String: sql = "INSERT INTO " & kvsName & sfx_old & " SELECT *, '{0}' AS SysEndTime FROM " & kvsName & " WHERE k='{1}'"
    Dim rst As Recordset: Set rst = dbLocal.OpenRecordset(kvsName, dbOpenDynaset)
    With rst
        .FindFirst "k='" & s_Key & "'"
        ts = Now
        If Not .NoMatch Then
            sql = Replace(sql, "{0}", SqlDateTime(ts))
            sql = Replace(sql, "{1}", rst!k)
            dbLocal.Execute sql                                 ' dbFailOnError not needed, temporal granularity (i.e. quietly denies write)
            .Delete
        End If
    End With
    rst.Close
    Set rst = Nothing
End Sub

Public Sub App_ValueSet(s_Key As String, s_Value As String)
    If s_Value = vbNullString Then Exit Sub
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
            dbLocal.Execute sql                               ' dbFailOnError not needed, temporal granularity (e.g. quietly denies > 1 append per minute)
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
End Sub

Public Function Hasher_FromFile(src As String) As String
    
    Dim trg As String: trg = User_TempFiles & "hasher.dat"
    
    Dim cmd As String: cmd = App_Assets & "hasher.exe --f ""{0}"" ""{1}"""
    cmd = Replace(cmd, "{0}", src)
    cmd = Replace(cmd, "{1}", trg)
    
    With New WshShell
        .Run cmd, 0, True
    End With
    
    With New FileSystemObject
        If .FileExists(trg) Then
            Hasher_FromFile = Trim(Replace(.OpenTextFile(trg).ReadAll(), vbCrLf, ""))
            .DeleteFile trg
        End If
    End With
    
End Function

Public Function Hasher_FromString(src As String) As String
    
    Dim trg As String: trg = User_TempFiles & "hasher.dat"
    
    Dim cmd As String: cmd = App_Assets & "hasher.exe --s ""{0}"" ""{1}"""
    cmd = Replace(cmd, "{0}", src)
    cmd = Replace(cmd, "{1}", trg)
    
    With New WshShell
        .Run cmd, 0, True
    End With
    
    With New FileSystemObject
        If .FileExists(trg) Then
            Hasher_FromString = Trim(Replace(.OpenTextFile(trg).ReadAll(), vbCrLf, ""))
            .DeleteFile trg
        End If
    End With
    
End Function

Public Function MailParser_GetJson(src As String) As String
    
    Dim trg As String: trg = User_TempFiles & "mail-parser.json"
    
    Dim cmd As String: cmd = App_Assets & "mail-parser.exe ""{0}"" ""{1}"""
    cmd = Replace(cmd, "{0}", src)
    cmd = Replace(cmd, "{1}", trg)
    
    With New WshShell
        .Run cmd, 0, True
    End With
    
    With New FileSystemObject
        If .FileExists(trg) Then
            MailParser_GetJson = Trim(.OpenTextFile(trg).ReadAll())
            .DeleteFile trg
        End If
    End With
    
End Function

Public Property Get Model_Current() As String
    Model_Current = s_Model
End Property

Public Property Let Model_Current(value As String)
    s_Model = value
End Property

Public Function Models_MRU() As Dictionary

    Dim d As Dictionary: Set d = New Dictionary
    
    Dim sql As String: sql = "SELECT v, Max(SysStartTime) As UsedDate FROM " & kvsName & sfx_old & " GROUP BY k, v HAVING k = '{0}' UNION SELECT v, SysStartTime As UsedDate FROM " & kvsName & " WHERE k = '{0}' ORDER BY UsedDate DESC;"
    sql = Replace(sql, "{0}", "modelDb")
   
    Dim rst As Recordset: Set rst = dbLocal.OpenRecordset(sql, dbOpenForwardOnly)
    Dim strDisplay, strDate As String
    Do Until rst.EOF
        If Not d.Exists(rst.Fields("v").value) Then
            strDate = mod_DateUtils.ToSSNNHHDDMMYYYY(rst.Fields("UsedDate").value)
            strDisplay = rst.Fields("v").value & ";" & mod_FsUtils.GetBaseName(rst.Fields("v").value) & ";" & strDate
            d.Add rst.Fields("v").value, strDisplay
        End If
        rst.MoveNext
    Loop
    rst.Close
    Set rst = Nothing
    
    Set Models_MRU = d

End Function

Public Sub Models_MRU_Clear()
    dbLocal.Execute "DELETE * FROM " & kvsName & sfx_old & " WHERE k = '" & "modelDb';"
End Sub

Public Sub Models_MRU_RemoveSelected(strPath As String)
    dbLocal.Execute "DELETE * FROM " & kvsName & sfx_old & " WHERE k = '" & "modelDb' AND v = '" & strPath & "';"
End Sub

Public Sub StatusBar_Clear()
    SysCmd acSysCmdClearStatus
End Sub

Public Sub StatusBar_Set(message As Variant)
    SysCmd acSysCmdSetStatus, message
End Sub

Public Property Get User_Email() As String
    User_Email = App_ValueGet("user-email")
End Property

Public Property Get User_Identity() As String
    User_Identity = App_ValueGet("user-id")
    If User_Identity = vbNullString Then User_Identity = vbNullUser
End Property

Public Function User_IsKnown() As Boolean
    User_IsKnown = False
    Select Case True
        Case App_Mode <> RemoteBackend: Exit Function
        Case User_Identity = vbNullUser: Exit Function
        Case User_Email = vbNullString: Exit Function
        Case User_Name = vbNullString: Exit Function
        Case Else: User_IsKnown = True
    End Select
End Function

Public Property Get User_Name() As String
    User_Name = App_ValueGet("user-name")
End Property

Public Property Get User_TempFiles() As String
    User_TempFiles = Environ$("LOCALAPPDATA") & vbBackSlash & "FMMS" & vbBackSlash
End Property

