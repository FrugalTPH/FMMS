Attribute VB_Name = "mod_Crypto"
Option Explicit

Private Type GUID_TYPE
    Data1 As Long
    Data2 As Long
    Data3 As Long
    Data4(7) As Byte
End Type

#If VBA7 Then
    Private Declare PtrSafe Function CoCreateGuid Lib "ole32.dll" (guid As GUID_TYPE) As LongPtr
    Private Declare PtrSafe Function StringFromGUID2 Lib "ole32.dll" (guid As GUID_TYPE, ByVal lpStrGuid As LongPtr, ByVal cbMax As Long) As LongPtr
#Else
    Private Declare Function CoCreateGuid Lib "ole32.dll" (guid As GUID_TYPE) As LongPtr
    Private Declare Function StringFromGUID2 Lib "ole32.dll" (guid As GUID_TYPE, ByVal lpStrGuid As LongPtr, ByVal cbMax As Long) As LongPtr
#End If


Public Function CreateGuid() As String
    Dim guid As GUID_TYPE
    Dim strGuid As String
    Dim retValue As LongPtr
    Const guidLength As Long = 39
    retValue = CoCreateGuid(guid)
    If retValue = 0 Then
        strGuid = String$(guidLength, vbNullChar)
        retValue = StringFromGUID2(guid, StrPtr(strGuid), guidLength)
        If retValue = guidLength Then CreateGuid = mod_StringUtils.TrimNulls(strGuid)
    End If
End Function
