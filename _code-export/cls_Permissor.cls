VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cls_Permissor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private frm As Form
Private lng_dir As Long
Private lng_scheme As Long
Private str_object As String
Private bln_isAttached As Boolean
Private bln_isBlocked As Boolean
Private bln_isCurUsr As Boolean
Private bln_isDeleted As Boolean
Private bln_isFrozen As Boolean
Private bln_IsFsConnected As Boolean
Private bln_isIdentified As Boolean
Private bln_isMgr As Boolean
Private bln_IsReadOnly As Boolean
Private bln_isSchemeMember As Boolean


Public Sub Init(F As Form)
On Error Resume Next
    Set frm = F
    lng_dir = Nz(frm!dir, 0)
    str_object = Nz(frm!object, vbNullString)
    Refresh
End Sub

Private Sub Class_Terminate()
    Set frm = Nothing
End Sub

Public Sub Refresh()
    lng_scheme = Nz(Form_fsub_Schemes!dir, 0)
    

    bln_isCurUsr = "P" & g_Model.User_Current = str_object
    bln_isMgr = g_Model.User_IsManager
    bln_IsReadOnly = g_Model.Db_isReadOnly
    bln_isSchemeMember = InStr(g_Model.User_CurrentSchemes, "," & lng_scheme & ",") > 0
    bln_IsFsConnected = g_Model.Fs_isConnected
    
On Error Resume Next
    bln_isAttached = InStr(Nz(frm!attachedTo, vbNullString), "," & lng_scheme & ",") > 0
    bln_isBlocked = Nz(frm!isBlocked, False)
    bln_isDeleted = Nz(frm!sysEndPerson, 0) > 0
    bln_isFrozen = Nz(frm!isFrozen, False)
    bln_isIdentified = Nz(frm!identity, vbNullString) <> vbNullString
End Sub

Public Property Get Can_Attach() As Boolean
    Can_Attach = (bln_isMgr Or bln_isSchemeMember) And Not bln_isDeleted And Not bln_isAttached
End Property

Public Property Get Can_Detach() As Boolean
    Can_Detach = (bln_isMgr Or bln_isSchemeMember) And Not bln_isDeleted And bln_isAttached
End Property

Public Property Get Can_Duplicate() As Boolean
    Can_Duplicate = bln_IsFsConnected And Not bln_isDeleted
End Property

Public Property Get Can_EditFs() As Boolean
    Can_EditFs = bln_IsFsConnected And Not bln_IsReadOnly
End Property

Public Property Get Can_EditModel() As Boolean
    Can_EditModel = Not bln_IsReadOnly
End Property

Public Property Get Can_EditUncertainty() As Boolean
    Can_EditUncertainty = Not (bln_IsReadOnly Or bln_isDeleted Or bln_isFrozen)
End Property

Public Property Get Can_Freeze() As Boolean
    Can_Freeze = bln_isMgr And Not bln_isFrozen And Not bln_isDeleted
End Property

Public Property Get Can_Manage() As Boolean
    Can_Manage = bln_isMgr
End Property

Public Property Get Can_ReadFs() As Boolean
    Can_ReadFs = bln_IsFsConnected
End Property

Public Property Get Can_RenumberWBS() As Boolean
    Can_RenumberWBS = bln_isMgr And Not bln_isDeleted
End Property

Public Property Get Can_Revise() As Boolean
    Can_Revise = bln_IsFsConnected And Not bln_isDeleted
End Property

Public Property Get Can_TakeSnapshot() As Boolean
    Can_TakeSnapshot = bln_IsFsConnected And Not (bln_IsReadOnly Or bln_isDeleted Or bln_isFrozen)
End Property

Public Property Get Can_Undelete() As Boolean
    Can_Undelete = bln_isMgr And bln_isDeleted
End Property

Public Property Get Can_Unfreeze() As Boolean
    Can_Unfreeze = bln_isMgr And bln_isFrozen And Not bln_isDeleted
End Property

Public Property Get Can_UpdatePermissions() As Boolean
    Can_UpdatePermissions = bln_isMgr And Not bln_isFrozen And Not bln_isDeleted
End Property

Public Property Get Can_Block() As Boolean
    Can_Block = bln_isMgr And bln_isIdentified And Not (bln_isBlocked Or bln_isCurUsr Or bln_isDeleted Or bln_isFrozen)
End Property

Public Property Get Can_Unblock() As Boolean
    Can_Unblock = bln_isMgr And bln_isIdentified And bln_isBlocked And Not (bln_isCurUsr Or bln_isDeleted Or bln_isFrozen)
End Property

Public Property Get scheme() As String
    scheme = "S" & lng_scheme
End Property
