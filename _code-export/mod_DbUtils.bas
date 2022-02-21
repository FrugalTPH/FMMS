Attribute VB_Name = "mod_DbUtils"
Option Explicit


Public Function dbLocal(Optional bolCleanup As Boolean = False) As DAO.Database
On Error GoTo errHandler
    Static dbCurrent As DAO.Database
    If bolCleanup Then GoTo CloseDb

retryDB:
    If dbCurrent Is Nothing Then Set dbCurrent = CurrentDb()
    Dim strTest As String: strTest = dbCurrent.name

exitRoutine:
    Set dbLocal = dbCurrent
    Exit Function

CloseDb:
    If Not (dbCurrent Is Nothing) Then Set dbCurrent = Nothing
    GoTo exitRoutine

errHandler:
    Select Case Err.Number
        Case 3420
            Set dbCurrent = Nothing
            If Not bolCleanup Then
                Resume retryDB
            Else
                Resume CloseDb
            End If
        Case Else
            MsgBox Err.Number & ": " & Err.Description, vbExclamation, "Error in dbLocal()"
            Resume exitRoutine
    End Select
End Function

Public Function ECount(domain As String, Optional criteria As String) As Long
On Error GoTo Err_ECount
    ECount = 0
    Dim strSql As String: strSql = "SELECT COUNT(*) AS ECount FROM " & domain
    If criteria <> vbNullString Then strSql = strSql & " WHERE " & criteria
    strSql = strSql & ";"
    
    Dim rst As Recordset: Set rst = dbLocal.OpenRecordset(strSql, dbOpenForwardOnly)
    ECount = rst!ECount
    
    rst.Close

Exit_ECount:
    Set rst = Nothing
    Exit Function

Err_ECount:
    MsgBox Err.Description, vbExclamation, "ECount Error " & Err.Number
    #If DEBUGMODE Then
        Stop
    #End If
    Resume Exit_ECount
End Function

Public Function ELookup(Expr As String, domain As String, Optional criteria As Variant = vbNullString, Optional OrderClause As Variant = vbNullString) As Variant
On Error GoTo Err_ELookup

    Dim varResult As Variant: varResult = Null
    Dim strSql As String: strSql = "SELECT TOP 1 " & Expr & " FROM " & domain
    If criteria <> vbNullString Then strSql = strSql & " WHERE " & criteria
    If OrderClause <> vbNullString Then strSql = strSql & " ORDER BY " & OrderClause
    strSql = strSql & ";"
    
    Dim rst As Recordset: Set rst = dbLocal.OpenRecordset(strSql, dbOpenForwardOnly)
    If rst.RecordCount > 0 Then varResult = rst(0)
    rst.Close
    
    ELookup = varResult

Exit_ELookup:
    Set rst = Nothing
    Exit Function

Err_ELookup:
    '#If DEBUGMODE Then
        'MsgBox Err.Description, vbExclamation, "ELookup Error " & Err.Number
        'Stop
    '#End If
    Resume Exit_ELookup
End Function

Public Sub Query_Refresh(qryName As String, sql As String)
    Dim q As DAO.QueryDef: Set q = dbLocal.QueryDefs(qryName)
    If q Is Nothing Then
        Set q = dbLocal.CreateQueryDef(qryName)
        dbLocal.QueryDefs.Refresh
    End If
    q.sql = sql
    Set q = Nothing
End Sub

Public Sub Table_Delete(tbl As String)
On Error Resume Next
    DoCmd.DeleteObject acTable, tbl
End Sub

Public Function Table_Exists(db As Database, tblName As String) As Boolean
On Error Resume Next
    Table_Exists = False
    Table_Exists = IsObject(db.TableDefs(tblName))
End Function

Public Sub WriteBytesToFile(vdata As Variant, trg As String)
    With New ADODB.stream
        .Type = adTypeBinary
        .Open
        .Write vdata
        .SaveToFile trg, adSaveCreateOverWrite
        .Close
    End With
End Sub

Public Sub WriteFirstAttachmentToFile(v_content As Variant, fpath As String, Optional rpath As String = vbNullString)

    Dim att As Recordset
    Dim fld As Field2
    
    Set att = v_content
        If Not (att.EOF And att.BOF) Then
            att.MoveFirst
            Set fld = att.Fields("FileData")
            If File_Exists(fpath) Then File_Delete fpath, True
            fld.SaveToFile fpath
            Set fld = Nothing
        End If
        att.Close
    Set att = Nothing
    
    If rpath <> vbNullString Then File_Rename fpath, rpath
    
End Sub
