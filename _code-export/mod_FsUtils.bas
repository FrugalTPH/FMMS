Attribute VB_Name = "mod_FsUtils"
Option Explicit

Public Enum copyOption
    coSkip = 0
    coMutate = 1
    coOverwrite = 2
End Enum


Public Sub Execute(strPath As String)
    If File_Exists(strPath) Or Folder_Exists(strPath) Then
        Dim strArgument As String: strArgument = "explorer.exe """ & strPath & """"
        Shell strArgument, vbNormalFocus
    Else
        MsgBox "Directory not found", vbExclamation, "Error"
    End If
End Sub

Public Sub File_Copy(src As String, trg As String, makeReadOnly As Boolean, copyOption As copyOption)
    With New FileSystemObject
        If .FileExists(src) Then
            If .FileExists(trg) Then
                Select Case copyOption
                    Case coOverwrite:
                    Case coMutate: trg = File_MutateName(trg)
                    Case Else: Exit Sub
                End Select
            End If
            .CopyFile src, trg, True
            If makeReadOnly Then File_SetReadOnly trg Else File_SetReadWrite trg
        End If
    End With
End Sub

Public Sub File_Delete(src As String, force As Boolean)
On Error Resume Next
    With New FileSystemObject
        If .FileExists(src) Then .DeleteFile src, force
    End With
End Sub

Public Function File_Exists(strPath As String) As Boolean
    With New FileSystemObject
        File_Exists = .FileExists(strPath)
    End With
End Function

'Public Sub File_Move(src As String, trg As String, makeReadOnly As Boolean)
'    With New FilesystemObject
'       If .FileExists(src) Then
'           If .FileExists(trg) Then If MsgBox("Do you wish to overwrite the existing file at: " & vbCrLf & trg, vbOKCancel, "Overwrite File") <> vbOK Then Exit Sub
'           .MoveFile src, trg
'           If makeReadOnly Then File_SetReadOnly trg Else File_SetReadWrite trg
'           File_Delete src, true
'       End If
'    End With
'End Sub

Public Function File_MutateName(src As String, Optional suffix As String = vbNullString) As String
    If suffix = vbNullString Then suffix = CStr(Date2Long)
    With New FileSystemObject
        File_MutateName = TrimTrailingChr(.GetParentFolderName(src), vbBackSlash) & vbBackSlash & _
            .GetBaseName(src) & _
            "_" & suffix & _
            "." & .GetExtensionName(src)
    End With
End Function

Public Function File_ReadText(strPath As String) As String
    File_ReadText = vbNullString
    With New FileSystemObject
        If .FileExists(strPath) Then File_ReadText = .OpenTextFile(strPath, ForReading, False, TristateTrue).ReadAll()
    End With
End Function

Public Function File_Rename(src As String, trg As String) As String
    If File_Exists(src) And Not File_Exists(trg) Then
        Name src As trg
        File_Rename = trg
    End If
End Function

Public Sub File_SetReadOnly(strPath As String)
    SetAttr strPath, vbReadOnly
End Sub

Public Sub File_SetReadWrite(strPath As String)
    SetAttr strPath, vbNormal
End Sub

Public Function File_WriteText(srcText As String, trg As String)
    With New FileSystemObject
        Dim txtStream As TextStream: Set txtStream = .CreateTextFile(trg, True, True)
        txtStream.WriteLine srcText
        txtStream.Close
        Set txtStream = Nothing
    End With
End Function

Public Sub Folder_Copy(src As String, trg As String)
    With New FileSystemObject
        If .FolderExists(src) And Not .FolderExists(trg) Then .CopyFolder src, trg, True
    End With
End Sub

Function Folder_Exists(strPath As Variant) As Boolean
    With New FileSystemObject
        Folder_Exists = .FolderExists(strPath)
    End With
End Function

Public Sub Folder_SetReadOnly(strPath As String)
    Dim C As Collection: Set C = New Collection
    RecursiveDir strPath, C
    Dim F As Variant
    For Each F In C
        File_SetReadOnly CStr(F)
    Next
End Sub

Public Sub Folder_SetReadWrite(strPath As String)
    Dim C As Collection: Set C = New Collection
    RecursiveDir strPath, C
    Dim F As Variant
    For Each F In C
        File_SetReadWrite CStr(F)
    Next
End Sub

Public Function GetBaseName(strPath As String) As String
' strPath = File:       Returns filename only (exc. extension)
' strPath = Folder:     Returns foldername only
    With New FileSystemObject
        GetBaseName = .GetBaseName(strPath)
    End With
End Function

Public Function GetExtensionName(strPath As String) As String
' strPath = File:       Returns extension only (exc. full-stop)
' strPath = Folder:     Returns empty string
    With New FileSystemObject
        GetExtensionName = .GetExtensionName(strPath)
    End With
End Function

'Public Function GetFileSize(strPath As String) As Double
'    With New FileSystemObject
'        Dim F As File: Set F = .GetFile(strPath)
'        GetFileSize = Nz(F.Size, 0)
'        Set F = Nothing
'    End With
'End Function

Function GetFolderPath(strPath As String) As String
    With New FileSystemObject
        GetFolderPath = .GetParentFolderName(strPath)
        If GetFolderPath <> vbNullString Then GetFolderPath = GetFolderPath & vbBackSlash
    End With
End Function

Public Function MkDirTree(sPath As String) As String
    If File_Exists(sPath) Or Folder_Exists(sPath) Or Nz(sPath, vbNullString) = vbNullString Then GoTo MyExit
    Dim aDirs As Variant: aDirs = Split(sPath, vbBackSlash)
    Dim iStart As Long: iStart = 1
    If Left$(sPath, 2) = vbBackSlash & vbBackSlash Then iStart = 3
    Dim sCurDir As String: sCurDir = Left$(sPath, InStr(iStart, sPath, vbBackSlash))
    Dim i As Long
    For i = iStart To UBound(aDirs)
        sCurDir = sCurDir & aDirs(i) & vbBackSlash
        If dir(sCurDir, vbDirectory) = vbNullString Then MkDir sCurDir
    Next i
MyExit:
    MkDirTree = sPath
End Function

Public Function RecursiveDir(strFolder As String, Optional colFiles As Collection = Null) As Collection

    If colFiles Is Nothing Then Set colFiles = New Collection
    strFolder = TrailingSlash(strFolder)
    Dim strTemp As String: strTemp = dir(strFolder & "*.*")                         ' Add all files in strFolder to colFiles
    Do While strTemp <> vbNullString
        colFiles.Add strFolder & strTemp
       strTemp = dir
    Loop
    Dim colFolders As New Collection                                                ' Fill colFolders with any found subdirectories
    strTemp = dir(strFolder, vbDirectory)
    
    Do While strTemp <> vbNullString
        If (strTemp <> ".") And (strTemp <> "..") Then
            If (GetAttr(strFolder & strTemp) And vbDirectory) <> 0 Then colFolders.Add strTemp
        End If
        strTemp = dir
    Loop
    
    Dim vFolderName As Variant                                                      ' Call same function for every subdirectory
    For Each vFolderName In colFolders
        RecursiveDir strFolder & vFolderName, colFiles
    Next vFolderName
    Set colFolders = Nothing
    Set RecursiveDir = colFiles
    
End Function
