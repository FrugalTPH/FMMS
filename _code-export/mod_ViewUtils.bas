Attribute VB_Name = "mod_ViewUtils"
Option Explicit

Public Const faceId_Attach As Long = 2308              ' 1087, 1079
Public Const faceId_Block As Long = 9643               ' 9934
Public Const faceId_Unblock As Long = 9642             ' 9934
Public Const faceId_Detach As Long = 2309              ' 1088
Public Const faceId_Comments As Long = 201             ' 1594
Public Const faceId_DeliveryPlan As Long = 9325        ' 11365
Public Const faceId_Duplicate As Long = 19             '
Public Const faceId_Freeze As Long = 6695              ' 1663
Public Const faceId_Unfreeze As Long = 6696            '
Public Const faceId_Permissions As Long = 7245         ' 7245, 9571
Public Const faceId_WbsRenumber As Long = 5882         ' 6100, 9466, 9453
Public Const faceId_Revise As Long = 2646              ' 525, 743, 2646, 7641
Public Const faceId_SnapshotAll As Long = 9304         '
Public Const faceId_SnapshotLatest As Long = 9305      '
Public Const faceId_SnapshotSave As Long = 20113       ' 280, 18673, 12867
Public Const faceId_Uncertainty As Long = 463          ' 3627, 3998, 5602, 6024
Public Const faceId_Undelete As Long = 4353            ' 680, 2552, 4353
Public Const faceId_WorkInProgress As Long = 7167      ' 1669

Private Const WM_SETICON = &H80
Private Const IMAGE_ICON = 1
Private Const LR_LOADFROMFILE = &H10
Private Const SM_CXSMICON As Long = 49
Private Const SM_CYSMICON As Long = 50

#If VBA7 Then
    Private Declare PtrSafe Function LoadImage Lib "user32" Alias "LoadImageA" (ByVal hInst As Long, ByVal lpsz As String, ByVal un1 As Long, ByVal n1 As Long, ByVal n2 As Long, ByVal un2 As Long) As Long
    Private Declare PtrSafe Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
    Private Declare PtrSafe Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
#Else
    Private Declare Function LoadImage Lib "user32" Alias "LoadImageA" (ByVal hInst As Long, ByVal lpsz As String, ByVal un1 As Long, ByVal n1 As Long, ByVal n2 As Long, ByVal un2 As Long) As Long
    Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
    Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
#End If


Public Function AddCustomButton(cbar As CommandBar, strCaption As String, strAction As String, Optional lng_faceId As Long = 0) As CommandBarButton
    
    Dim btn As CommandBarButton: Set btn = cbar.Controls.Add(, , , , True)
    With btn
        .OnAction = strAction
        .Caption = strCaption
        .Style = msoButtonIconAndCaptionBelow
        .height = 24
        .Width = 24
        If lng_faceId > 0 Then .FaceId = lng_faceId
    End With
    
    Set AddCustomButton = btn
    
End Function

Public Function AddCustomSubButton(cpu As CommandBarPopup, strCaption As String, strAction As String, Optional lng_faceId As Long = 0) As CommandBarButton
    
    Dim btn As CommandBarButton: Set btn = cpu.Controls.Add(, , , , True)
    With btn
        .OnAction = strAction
        .Caption = strCaption
        .Style = msoButtonIconAndCaptionBelow
        .height = 24
        .Width = 24
        If lng_faceId > 0 Then .FaceId = lng_faceId
    End With
    
    Set AddCustomSubButton = btn
    
End Function

Public Function CloseOpenForms(Optional strExcept As String)
    Dim i As Integer
    With Application.Forms
        For i = .count - 1 To 0 Step -1
           With .Item(i)
               If .name <> strExcept Then DoCmd.Close acForm, .name
           End With
       Next i
    End With
End Function

Public Sub ToggleForm(strForm As String)
    If CurrentProject.AllForms(strForm).IsLoaded Then
        DoCmd.Close acForm, strForm, acSaveNo
    Else
        DoCmd.OpenForm strForm
    End If
End Sub

Public Function TryGetClassification(fldName As String, prefix As String) As String
On Error GoTo eh
    TryGetClassification = vbNullString
    Dim count As Long: count = ECount("tbl_Definitions_cache", "field='" & fldName & "' AND " & prefix & ">0")
    Select Case count
        Case 1:         TryGetClassification = ELookup("code", "tbl_Definitions_cache", "field='" & fldName & "' AND " & prefix & ">0")
        Case Is > 1:    TryGetClassification = ELookup("code", "tbl_Definitions_cache", "field='" & fldName & "' AND " & prefix & ">0", prefix & " DESC, sort, code")
    End Select
eh:
End Function

Public Function TryGetSchemeValue(fldName As String, prefix As String, Optional scheme As Long = 0) As Variant
On Error GoTo eh
    If scheme = 0 Then
        TryGetSchemeValue = GetForm("fsub_Schemes").Controls(fldName)
    Else
        TryGetSchemeValue = ELookup(fldName, "quni_Schemes", "dir=" & scheme)
    End If
eh:
End Function

Public Function GetForm(formName As String) As Form
On Error Resume Next
    Select Case formName
        Case "fsub_Schemes", "S": Set GetForm = Forms("frm_Schemes").sub_Browser.Form
        Case "fsub_People", "P": Set GetForm = Forms("frm_People").sub_Browser.Form
        Case "fsub_Organisations", "O": Set GetForm = Forms("frm_Organisations").sub_Browser.Form
        Case "fsub_Inputs", "N": Set GetForm = Forms("frm_Inputs").sub_Browser.Form
        Case "fsub_Emails", "E": Set GetForm = Forms("frm_Emails").sub_Browser.Form
        Case "fsub_Outputs", "U": Set GetForm = Forms("frm_Outputs").sub_Browser.Form
        Case "fsub_Memoranda", "M": Set GetForm = Forms("frm_Memoranda").sub_Browser.Form
        Case "fsub_Filestage": Set GetForm = Forms("frm_Filestage").sub_Browser.Form
        Case "fsub_Mailstage": Set GetForm = Forms("frm_Mailstage").sub_Browser.Form
        Case Else: Set GetForm = Forms(formName)
    End Select
End Function

Public Sub Limit_FormSize(frm As Form)
    With frm
        If .InsideWidth < .w Then .InsideWidth = .w
        If .InsideHeight < .h Then .InsideHeight = .h
    End With
End Sub

Public Sub Set_FormIcon(frm As Form, strIconName As String)
    Dim X As Long: X = GetSystemMetrics(SM_CXSMICON)
    Dim Y As Long: Y = GetSystemMetrics(SM_CYSMICON)
    Dim lIcon As Long: lIcon = LoadImage(0, g_App.App_Assets & strIconName & ".ico", 1, X, Y, LR_LOADFROMFILE)
    Dim lResult As Long: lResult = SendMessage(frm.hwnd, WM_SETICON, 0, ByVal lIcon)
End Sub

Public Sub Set_FormSize(frm As Form)
    With frm
        Select Case frm.name
            Case "fdlg_Classifier":             .w = 3500:  .h = 6750
            Case "fdlg_Commenter":              .w = 14000: .h = 3500
            Case "fdlg_Filestage_matches":      .w = 18000: .h = 2700
            Case "fdlg_Login":                  .w = 7000:  .h = 10600
            Case "fdlg_Mailstage_matches":      .w = 18000: .h = 2700
            Case "fdlg_NewModel":               .w = 0:     .h = 0
            Case "fdlg_Permissor":              .w = 8460:  .h = 3600
            Case "fdlg_Templater":              .w = 20000: .h = 3500
            Case "fdlg_Uncertainty":            .w = 6350:  .h = 4080
            Case "fdlg_WordPad":                .w = 14000: .h = 3500
            Case "frm_Definitions":             .w = 17000: .h = 2700
            Case "frm_Emails":                  .w = 17000: .h = 2700
            Case "frm_Filestage":               .w = 20000: .h = 3500
            Case "frm_Home":                    .w = 15000: .h = 6000
            Case "frm_Inputs":                  .w = 17000: .h = 2700
            Case "frm_Mailstage":               .w = 20000: .h = 3500
            Case "frm_Memoranda":               .w = 17000: .h = 2700
            Case "frm_Organisations":           .w = 17000: .h = 2700
            Case "frm_Outputs":                 .w = 17000: .h = 2700
            Case "frm_People":                  .w = 17000: .h = 2700
            Case "frm_Schemes":                 .w = 17000: .h = 2850
            Case Else:                          .w = 0:     .h = 0
        End Select
    End With
End Sub

Public Sub Set_FormPermissions(trg As Form, permSrc As Form)
    With New cls_Permissor
        .Init permSrc
        Dim ctl As Control
        For Each ctl In trg.Controls
            ctl.Visible = True
            Select Case True
                Case InStr(ctl.Tag, "rw") > 0: If Not .Can_EditModel Then ctl.Visible = False
                Case InStr(ctl.Tag, "fs") > 0: If Not .Can_EditFs Then ctl.Visible = False
                Case InStr(ctl.Tag, "mgr") > 0: If Not .Can_Manage Then ctl.Visible = False
            End Select
        Next ctl
        Set ctl = trg!sub_Browser
        Dim frm As Form: Set frm = ctl.Form
        frm.AllowAdditions = .Can_EditModel
        frm.AllowDeletions = .Can_EditModel
        frm.AllowEdits = .Can_EditModel
    End With
End Sub
