Attribute VB_Name = "std_Messages"
Option Explicit


Public Const Filestage_FileDropWarning1 = "One or more invalid filetypes were detected and will be skipped." & vbCrLf & _
    "This could include any of the following:" & vbCrLf & vbCrLf & _
    "Type:" & vbTab & vbTab & "Resolution:" & vbCrLf & _
    ".url" & vbTab & "=>" & vbTab & "Use the Filestage (paste URL)" & vbCrLf & _
    ".msg" & vbTab & "=>" & vbTab & "Use the Mailstage" & vbCrLf & _
    ".eml" & vbTab & "=>" & vbTab & "Use the Mailstage" & vbCrLf & _
    ".lnk" & vbTab & "=>" & vbTab & "Bad practice (capture the target file instead)"

Public Const Mailstage_FileDropWarning1 = "One or more invalid filetypes were detected and will be skipped." & vbCrLf & _
    "Only the following email formats are currently supported:" & vbCrLf & vbCrLf & _
    ".eml" & vbTab & "=>" & vbTab & "MIME RFC 822 Standard" & vbCrLf & _
    ".msg" & vbTab & "=>" & vbTab & "MS Exchange / Outlook" & vbCrLf

