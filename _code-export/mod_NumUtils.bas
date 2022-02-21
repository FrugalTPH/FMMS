Attribute VB_Name = "mod_NumUtils"
Option Explicit


Public Function ArrayLen(arr As Variant) As Long
    ArrayLen = UBound(arr) - LBound(arr) + 1
End Function

Public Function String2Long(ByVal strNum As String) As Long
    String2Long = 0
    If IsNumeric(strNum) Then String2Long = CLng(strNum)
End Function

Public Function CsvAdd_numSort(csv As String, num As Long, trimmed As Boolean) As String
    CsvAdd_numSort = csv
    If Nz(num, 0) <= 0 Then Exit Function
    Dim L As Long
    Dim v As Variant
    Dim a() As String: a = Split(csv, ",")
    Dim alist As Object: Set alist = CreateObject("System.Collections.ArrayList")
    alist.Add num
    For Each v In a
        L = String2Long(v)
        If Not alist.contains(L) And L > 0 And L <> num Then alist.Add L
    Next
    alist.sort
    CsvAdd_numSort = Join(alist.ToArray, ",")
    If CsvAdd_numSort = vbNullString Or trimmed Then Exit Function
    CsvAdd_numSort = "," & CsvAdd_numSort & ","
End Function

Public Function CsvRemove_numSort(csv As String, num As Long, trimmed As Boolean) As String
    CsvRemove_numSort = csv
    If Nz(num, 0) <= 0 Then Exit Function
    Dim L As Long
    Dim v As Variant
    Dim a() As String: a = Split(csv, ",")
    Dim alist As Object: Set alist = CreateObject("System.Collections.ArrayList")
    For Each v In a
        L = String2Long(v)
        If Not alist.contains(L) And L > 0 And L <> num Then alist.Add L
    Next
    alist.sort
    CsvRemove_numSort = Join(alist.ToArray, ",")
    If CsvRemove_numSort = vbNullString Or trimmed Then Exit Function
    CsvRemove_numSort = "," & CsvRemove_numSort & ","
End Function
