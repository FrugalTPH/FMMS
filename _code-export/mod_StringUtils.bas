Attribute VB_Name = "mod_StringUtils"
Option Explicit


Public Function CollapseRepeatedChar(str As String, chr As String) As String
    Dim lastLen As Long
    CollapseRepeatedChar = str
    Do Until Len(CollapseRepeatedChar) = lastLen
        lastLen = Len(CollapseRepeatedChar)
        CollapseRepeatedChar = Replace(CollapseRepeatedChar, "##", "#")
    Loop
End Function

Public Function CsvAdd_alphaSort(csv As String, val As Variant, trimmed As Boolean) As String
    CsvAdd_alphaSort = csv
    Dim alist As Object: Set alist = CreateObject("System.Collections.ArrayList")
    alist.Add CStr(val)
    Dim a() As String: a = Split(csv, ",")
    Dim v As Variant
    For Each v In a
        If Not alist.contains(v) And v <> val Then alist.Add v
    Next
    alist.sort
    CsvAdd_alphaSort = Join(alist.ToArray, ",")
    If CsvAdd_alphaSort = vbNullString Or trimmed Then Exit Function
    CsvAdd_alphaSort = "," & CsvAdd_alphaSort & ","
End Function

Public Function CsvRemove_alphaSort(csv As String, val As Variant, trimmed As Boolean) As String
    CsvRemove_alphaSort = csv
    Dim alist As Object: Set alist = CreateObject("System.Collections.ArrayList")
    Dim a() As String: a = Split(csv, ",")
    Dim v As Variant
    For Each v In a
        If Not alist.contains(v) And v <> CStr(val) Then alist.Add v
    Next
    alist.sort
    CsvRemove_alphaSort = Join(alist.ToArray, ",")
    If CsvRemove_alphaSort = vbNullString Or trimmed Then Exit Function
    CsvRemove_alphaSort = "," & CsvRemove_alphaSort & ","
End Function

Public Function EndsWith(str As String, ending As String) As Boolean
     Dim endingLen As Long: endingLen = Len(ending)
     EndsWith = (Right$(Trim$(UCase$(str)), endingLen) = UCase$(ending))
End Function

Public Function HyperlinkMidPart(hyperlink As String) As String
    Dim str As String: str = CollapseRepeatedChar(hyperlink, "#")
    If Len(str) - Len(Replace(str, "#", "")) <> 2 Then Exit Function
    Dim pLeft As Integer: pLeft = InStr(str, "#") + 1
    Dim pRight As Integer: pRight = InStrRev(str, "#")
    HyperlinkMidPart = Mid(str, pLeft, pRight - pLeft)
End Function

Public Function IsTruthy(str As String) As Boolean
    IsTruthy = False
    If str = vbNullString Then Exit Function
    If str = "true" Then GoTo truthy
    If str = "t" Then GoTo truthy
    If str = "yes" Then GoTo truthy
    If str = "y" Then GoTo truthy
    If val(str) > 0 Then GoTo truthy
    Exit Function
 
truthy:
    IsTruthy = True
End Function

'Public Function IsFalsey(str As String) As Boolean
'    IsFalsey = True
'    If str = vbNullString Then Exit Function
'    If str = "false" Then Exit Function
'    If str = "f" Then Exit Function
'    If str = "no" Then Exit Function
'    If str = "n" Then Exit Function
'    If str = "0" Then Exit Function
'    If val(str) < 0 Then Exit Function
'    IsFalsey = False
'End Function

Public Function KebabCase(strDirty As String) As String
    Dim strClean As String
    Dim regEx As Object: Set regEx = New RegExp
    With regEx
        .Global = True
        .MultiLine = True
        .IgnoreCase = False
        .Pattern = "[^a-zA-Z0-9]"
    End With
    strClean = regEx.Replace(strDirty, vbDash)                                              ' Convert disallowed characters
    
    Do While InStr(strClean, vbDash + vbDash) > 0                                           ' Remove duplicate delimiters
        strClean = Replace(strClean, vbDash + vbDash, vbDash)
    Loop
    If EndsWith(strClean, vbDash) Then strClean = Left(strClean, Len(strClean) - 1)         ' Remove leading & trailing delimiters
    If StartsWith(strClean, vbDash) Then strClean = Right(strClean, Len(strClean) - 1)
    KebabCase = LCase(strClean)                                                             ' Return in Lowercase
End Function

Public Function RightNumberPart(str As String, strDelimeter As String) As Long
    RightNumberPart = 0
    Dim lngPos As Long: lngPos = InStrRev(str, strDelimeter)
    Dim strPart As String: strPart = vbNullString
    If lngPos > 0 Then strPart = Mid(str, lngPos + 1)
    If Not strPart = vbNullString Then RightNumberPart = val(strPart)
End Function

Public Function Sanitize_Sql(ByVal sElementValue As String) As String
    Dim sValue As String: sValue = CStr(sElementValue)
    If Not IsNull(sValue) Then
        sValue = Replace(sValue, "'", "''")
        sValue = Replace(sValue, """", """""")
    End If
    Sanitize_Sql = sValue
End Function

Public Function StartsWith(str As String, start As String) As Boolean
     Dim startLen As Long: startLen = Len(start)
     StartsWith = (Left$(Trim$(UCase$(str)), startLen) = UCase$(start))
End Function

Public Function StripChrs(str As String, chr As String) As String
     StripChrs = Replace(str, chr, vbNullString)
End Function

Public Function StripNonAsciiChars(ByVal InputString As String) As String
    Dim regEx As New RegExp
    With regEx
        .Global = True
        .MultiLine = True
        .IgnoreCase = True
        .Pattern = "[^\u0000-\u007F]"
        StripNonAsciiChars = regEx.Replace(InputString, "")
    End With
End Function

Public Function StripToTheseCharacters(str As String, allowableChrs As String)
    StripToTheseCharacters = vbNullString
    Dim curChar As String
    Dim i As Integer
    For i = 1 To Len(str)
        curChar = Mid(str, i, 1)
        If InStr(UCase(allowableChrs), UCase(curChar)) Then StripToTheseCharacters = StripToTheseCharacters & curChar
    Next
End Function

Public Function TrailingSlash(varIn As Variant) As String
    If Len(varIn) > 0 Then
        If Right$(varIn, 1) = vbBackSlash Then
            TrailingSlash = varIn
        Else
            TrailingSlash = varIn & vbBackSlash
        End If
    End If
End Function

Public Function TrailingSlashRemove(s As String) As String
    Do While Right(s, 1) = vbBackSlash
        s = Left(s, Len(s) - 1)
    Loop
    TrailingSlashRemove = s
End Function

Public Function TreeIndent_Symbol(indent As Integer) As String
    If indent <= 1 Then
        TreeIndent_Symbol = String(indent * 2, ChrW(&HA0)) & ChrW(11206) & ChrW(&HA0)
    Else
        TreeIndent_Symbol = String(indent * 2, ChrW(&HA0)) & ChrW(11208) & ChrW(&HA0)
    End If
End Function

Public Function TreeIndent_Set(str As String, indent As Integer) As String
    str = StripNonAsciiChars(str)
    str = LTrim$(str)
    str = TreeIndent_Symbol(indent) & str
    TreeIndent_Set = str
End Function

Public Function TrimNulls(ByVal str As String) As String
    Dim iPos As Long: iPos = InStr(str, chr$(0))
    If iPos > 0 Then str = Left$(str, iPos - 1)
    TrimNulls = str
End Function

Public Function TrimLeadingChr(str As String, chr As String)
    TrimLeadingChr = str
    If Left(str, 1) = chr Then TrimLeadingChr = Right(str, Len(str) - 1)
End Function

Public Function TrimLeadingAndTrailingChr(str As String, chr As String)
    TrimLeadingAndTrailingChr = TrimTrailingChr(TrimLeadingChr(str, chr), chr)
End Function

Public Function TrimTrailingChr(str As String, chr As String)
    TrimTrailingChr = str
    If Right(str, 1) = chr Then TrimTrailingChr = Left(str, Len(str) - 1)
End Function

Public Function ValidDocRefNo(strDirty As String) As String
    Dim strClean As String
    Dim regEx As Object: Set regEx = New RegExp
    With regEx
        .Global = True
        .MultiLine = True
        .IgnoreCase = False
        .Pattern = "[^a-zA-Z0-9]"
    End With
    strClean = regEx.Replace(strDirty, vbDash)                  ' Replace disallowed characters with dashes
    Do While InStr(strClean, vbDash + vbDash) > 0               ' Collapse duplicate dashes
        strClean = Replace(strClean, vbDash + vbDash, vbDash)
    Loop
    strClean = TrimLeadingAndTrailingChr(strClean, vbDash)      ' Trim leading / trailing dashes
    ValidDocRefNo = UCase(strClean)                             ' Return in uppercase
End Function

Public Function ValidDocRevNo(strDirty As String) As String
    Dim strClean As String
    Dim regEx As Object: Set regEx = New RegExp
    With regEx
        .Global = True
        .MultiLine = True
        .IgnoreCase = False
        .Pattern = "[^a-zA-Z0-9]"
    End With
    strClean = regEx.Replace(strDirty, vbNullString)            ' Remove disallowed characters
    ValidDocRevNo = UCase(strClean)                             ' Return in uppercase
End Function

Public Function ValidUrl(sURL As String) As String
    ValidUrl = vbNullString
    If Len(sURL) <= 0 Then Exit Function
    If Not InStr(sURL, "http") > 0 Then sURL = "http://" & sURL
    Dim oXHTTP As Object: Set oXHTTP = CreateObject("MSXML2.XMLHTTP")
    oXHTTP.Open "HEAD", sURL, False
    On Error Resume Next
    oXHTTP.Send
    If oXHTTP.ReadyState = 4 Then
        If oXHTTP.status = 200 Then ValidUrl = sURL
        If oXHTTP.status = 0 Then ValidUrl = sURL
    End If
    Set oXHTTP = Nothing
End Function


