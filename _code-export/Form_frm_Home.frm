VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frm_Home"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public w As Long
Public h As Long
Private skipCloseSeq As Boolean


Private Sub btn_ModelBrowse_Click()
    Dim strPath As String
    With Application.FileDialog(3)
        .AllowMultiSelect = False
        .Filters.Add "Access Database", "*.accdb", 1
        .InitialFileName = g_App.App_Models
        .Title = "Select a valid FMMS model..."
        If .Show = -1 Then strPath = .SelectedItems.Item(1)
    End With
    g_Model.Db_Read strPath
    RefreshModelStatus
End Sub

Private Sub btn_ModelNew_Click()
On Error Resume Next
    DoCmd.OpenForm "fdlg_NewModel", , , , , acDialog
    RefreshModelStatus
End Sub

Private Sub btn_ModelOpen_Click()
    Model_Load lbo_Models
End Sub

Private Sub Form_Close()
    If skipCloseSeq Then Exit Sub
    g_App.Terminate
End Sub

Private Sub Form_Load()
    
    If g_App.App_Mode = LocalSnapshot Then
        Model_Load CurrentProject.FullName
        Exit Sub
    End If
    
    If g_App.App_Mode = RemoteBackend And g_App.User_IsKnown Then
        g_Model.Terminate                                       ' Clears any model remnants from previous bad termination
        RefreshModelStatus
    Else
    
        ' TODO: Identify & authenticate user - though for now, the programme just exits if app_Kvs isn't preset with user info (user-email, user-name & user-id).
        
        MsgBox "TODO: Identify & authenticate user - though for now, the programme just exits if app_Kvs isn't preset with user info (user-email, user-name & user-id).", vbCritical, "Error - frm_Home.Form_Load"
        
        skipCloseSeq = True
        DoCmd.Close acForm, Me.name, acSaveNo
        g_App.Terminate
        
    End If

End Sub

Private Sub Form_Open(Cancel As Integer)

    If g_App.App_Mode = Uninitialized Then g_App.Initialize

    skipCloseSeq = False
    Set_FormSize Me
    Set_FormIcon Me, "frugal"
    
End Sub

Private Sub Form_Resize()
    Limit_FormSize Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If skipCloseSeq Then Exit Sub
    Cancel = MsgBox("Ok to quit the application?", vbOKCancel, "Close Application") <> vbOK
End Sub

Private Sub lbl_ClearAll_Click()
    If MsgBox("Are you sure you want to clear your most recently used models list?", vbOKCancel, "Clear MRU") <> vbOK Then Exit Sub
    g_App.Models_MRU_Clear
    RefreshModelStatus
End Sub

Private Sub lbl_ClearSelected_Click()
    If MsgBox("Ok to remove the selected item from your most recently used models list?", vbOKCancel, "Remove MRU Item") <> vbOK Then Exit Sub
    g_App.Models_MRU_RemoveSelected lbo_Models
    RefreshModelStatus
End Sub

Private Sub lbo_Models_AfterUpdate()
    RefreshCaptions
End Sub

' ------- '
' PRIVATE '
' ------- '


Public Sub Model_Load(strPath As String)

    g_App.Model_Current = strPath
    
    g_Model.Initialize
    
    skipCloseSeq = True
    DoCmd.Close acForm, Me.name, acSaveNo
    DoCmd.OpenForm "frm_Schemes"
    
End Sub

Private Sub RefreshModelStatus()
    
    Dim bln_SignedIn As Boolean: bln_SignedIn = g_App.User_IsKnown
    
    lbo_Models.Visible = bln_SignedIn
    btn_ModelOpen.Visible = bln_SignedIn
    btn_ModelNew.Visible = bln_SignedIn
    btn_ModelBrowse.Visible = bln_SignedIn
    lbl_ClearAll.Visible = bln_SignedIn
    lbl_ClearSelected.Visible = bln_SignedIn
    
    If Not bln_SignedIn Then Caption = "Home" Else Caption = "Home   |   " & g_App.User_Name & "  (" & g_App.User_Email & ")"
    
    lbo_Models.RowSource = vbNullString
    
    With g_App.Models_MRU
        Dim k As Variant
        For Each k In .Keys()
            lbo_Models.AddItem .Item(k)
        Next
    End With
    
    If IsNull(lbo_Models.ItemData(0)) Or lbo_Models.ItemData(0) = vbNullString Then
        btn_ModelBrowse.SetFocus
        btn_ModelOpen.Visible = False
    Else
        lbo_Models = lbo_Models.ItemData(0)
        RefreshCaptions
    End If
    
End Sub

' ------ '
' PUBLIC '
' ------ '

Public Sub RefreshCaptions()
    Me.Caption = "Home   ~   " & g_App.User_Name & "  (" & g_App.User_Email & ")   ~   " & lbo_Models.Column(1)
End Sub


