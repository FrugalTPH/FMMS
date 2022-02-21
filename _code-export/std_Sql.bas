Attribute VB_Name = "std_Sql"
Option Explicit


Public Property Get cmb_Main_D() As String
    cmb_Main_D = "SELECT DISTINCT field FROM tbl_Definitions ORDER BY field;"
End Property

Public Property Get cmb_Main_E() As String
    cmb_Main_E = "All Emails;1;Scheme Attachments;2"
End Property

Public Property Get cmb_Main_M() As String
    cmb_Main_M = "All Memoranda;1;Scheme Attachments;2"
End Property

Public Property Get cmb_Main_M_Templates() As String
    cmb_Main_M_Templates = "SELECT code, title FROM tbl_Definitions_cache WHERE field='typeCode' AND m > 0 ORDER BY sort, code;"
End Property

Public Property Get cmb_Main_N() As String
    cmb_Main_N = "All Input Documents;1;Scheme Attachments;2"
End Property

Public Property Get cmb_Main_O() As String
    cmb_Main_O = "All Organisations;1;Scheme Attachees;2"
End Property

Public Property Get cmb_Main_P() As String
    cmb_Main_P = "All People;1;Scheme Attachees;2"
End Property

Public Property Get cmb_Main_S() As String
    cmb_Main_S = "SELECT title, code FROM tbl_Definitions_cache WHERE field = 'schemeCode' AND s > 0 ORDER BY sort, code;"
End Property

Public Property Get cmb_Main_U() As String
    cmb_Main_U = "All Output Documents;1;Scheme Attachments;2;Scheme Outputs;3"
End Property

Public Property Get cmb_Main_U_Templates() As String
    cmb_Main_U_Templates = "SELECT code, title FROM tbl_Definitions_cache WHERE field='typeCode' AND u > 0 ORDER BY sort, code;"
End Property

Public Property Get cmb_Main_W() As String
    cmb_Main_W = "SELECT quni_WordPad.sysStartTime, DLookUp('title','quni_People','quni_People.dir=' & [quni_WordPad].[sysStartPerson]) AS [user] " & _
                 "FROM quni_WordPad " & _
                 "WHERE (((quni_WordPad.object)='{0}'));" '& _
                 "ORDER BY quni_WordPad.sysStartTime DESC;"
End Property

Public Property Get Definitions_current() As String
    Definitions_current = "SELECT * FROM tbl_Definitions WHERE field = Forms![frm_Definitions]![cmb_Main] ORDER BY sort, code;"
End Property

Public Property Get Definitions_old() As String
    Definitions_old = "SELECT * FROM quni_Definitions AS uni " & _
                    "INNER JOIN ( SELECT code AS c, MAX(sysEndTime) AS max_et FROM quni_Definitions GROUP BY code ) AS grp " & _
                    "ON grp.max_et = uni.sysEndTime AND grp.c = uni.code " & _
                    "WHERE field = Forms![frm_Definitions]![cmb_Main] " & _
                    "ORDER BY uni.sysEndTime DESC, sort;"
End Property

Public Property Get Definitions_quni() As String
    Definitions_quni = "SELECT * FROM tbl_Definitions_old UNION ALL SELECT * FROM tbl_Definitions ORDER BY sysEndTime DESC , code DESC;"
End Property

Public Property Get Emails_current() As String
    Emails_current = "SELECT *, DMax(""sysStartTime"",""tbl__Comments"",""object='"" & [object] & ""'"") AS comment  FROM tbl_Emails ORDER BY sysStartTime DESC, dir DESC;"
End Property

Public Property Get Emails_old() As String
    Emails_old = "SELECT *, DMax(""sysStartTime"",""tbl__Comments"",""object='"" & [object] & ""'"") AS comment  FROM quni_Emails AS uni " & _
                "INNER JOIN ( SELECT dir AS d, MAX(sysEndTime) AS max_et FROM quni_Emails GROUP BY dir ) AS grp " & _
                "ON grp.max_et = uni.sysEndTime AND grp.d = uni.dir " & _
                "ORDER BY uni.sysEndTime DESC, uni.sysStartTime DESC;"
End Property

Public Property Get Emails_byCriteria() As String
    Emails_byCriteria = "SELECT *, DMax(""sysStartTime"",""tbl__Comments"",""object='"" & [object] & ""'"") AS comment  FROM quni_Emails AS uni " & _
                "INNER JOIN ( SELECT dir AS d, MAX(sysEndTime) AS max_et FROM quni_Emails GROUP BY dir ) AS grp " & _
                "ON grp.max_et = uni.sysEndTime AND grp.d = uni.dir " & vbCrLf & _
                "WHERE FALSE {0}" & vbCrLf & _
                "ORDER BY uni.sysEndTime DESC, uni.sysStartTime DESC;"
End Property

Public Property Get Emails_quni() As String
    Emails_quni = "SELECT * FROM tbl_Emails_old UNION ALL SELECT * FROM tbl_Emails ORDER BY sysEndTime DESC , dir DESC;"
End Property

Public Property Get fdlg_Filestage_matches(csv As String) As String
    fdlg_Filestage_matches = vbNullString
    If csv = vbNullString Then Exit Property
    fdlg_Filestage_matches = "SELECT * FROM quni_Inputs WHERE dir IN(" & csv & ");"
End Property

Public Property Get frm_Mailstage_matches(csv As String) As String
    frm_Mailstage_matches = vbNullString
    If csv = vbNullString Then Exit Property
    frm_Mailstage_matches = "SELECT * FROM quni_Emails WHERE dir IN(" & csv & ");"
End Property

Public Property Get fdlg_Permissor() As String
    fdlg_Permissor = "SELECT code, title FROM tbl_Definitions_cache WHERE field='roleCode' AND p > 0 ORDER BY code;"
End Property

Public Function fdlg_Classifier(formName As String, fieldName As String) As String
    Dim s_Type As String: s_Type = vbNullString
    Select Case formName
        Case "frm_Schemes", "fsub_Schemes": s_Type = "s"
        Case "frm_People", "fsub_People": s_Type = "p"
        Case "frm_Organisations", "fsub_Organisations": s_Type = "o"
        Case "frm_Emails", "fsub_Emails", "frm_Mailstage", "fsub_Mailstage": s_Type = "e"
        Case "frm_Inputs", "fsub_Inputs", "frm_Filestage", "fsub_Filestage": s_Type = "n"
        Case "frm_Memoranda", "fsub_Memoranda": s_Type = "m"
        Case "frm_Outputs", "fsub_Outputs": s_Type = "u"
        Case Else: Exit Function
    End Select
    fdlg_Classifier = "SELECT code, title FROM tbl_Definitions_cache WHERE field='" & fieldName & "' AND " & s_Type & " > 0 AND title ALike '%' & [Forms]![fdlg_Classifier]![txt_Query] & '%' ORDER BY code;"
End Function

Public Property Get Inputs_current() As String
    Inputs_current = "SELECT *, DMax(""sysStartTime"",""tbl__Comments"",""object='"" & [object] & ""'"") AS comment  FROM tbl_Inputs ORDER BY sysStartTime DESC, dir DESC;"
End Property

Public Property Get Inputs_old() As String
    Inputs_old = "SELECT *, DMax(""sysStartTime"",""tbl__Comments"",""object='"" & [object] & ""'"") AS comment  FROM quni_Inputs AS uni " & _
                "INNER JOIN ( SELECT dir AS d, MAX(sysEndTime) AS max_et FROM quni_Inputs GROUP BY dir ) AS grp " & _
                "ON grp.max_et = uni.sysEndTime AND grp.d = uni.dir " & _
                "ORDER BY uni.sysEndTime DESC, uni.sysStartTime DESC;"
End Property

Public Property Get Inputs_byCriteria() As String
    Inputs_byCriteria = "SELECT *, DMax(""sysStartTime"",""tbl__Comments"",""object='"" & [object] & ""'"") AS comment  FROM quni_Inputs AS uni " & _
                "INNER JOIN ( SELECT dir AS d, MAX(sysEndTime) AS max_et FROM quni_Inputs GROUP BY dir ) AS grp " & _
                "ON grp.max_et = uni.sysEndTime AND grp.d = uni.dir " & vbCrLf & _
                "WHERE FALSE {0}" & vbCrLf & _
                "ORDER BY uni.sysEndTime DESC, uni.sysStartTime DESC;"
End Property

Public Property Get Inputs_quni() As String
    Inputs_quni = "SELECT * FROM tbl_Inputs_old UNION ALL SELECT * FROM tbl_Inputs ORDER BY sysEndTime DESC , dir DESC;"
End Property

Public Property Get Memoranda_current() As String
    Memoranda_current = "SELECT *, DMax(""sysStartTime"",""tbl__Comments"",""object='"" & [object] & ""'"") AS comment  FROM tbl_Memoranda ORDER BY sysStartTime DESC, dir DESC;"
End Property

Public Property Get Memoranda_old() As String
    Memoranda_old = "SELECT *, DMax(""sysStartTime"",""tbl__Comments"",""object='"" & [object] & ""'"") AS comment  FROM quni_Memoranda AS uni " & _
                "INNER JOIN ( SELECT dir AS d, MAX(sysEndTime) AS max_et FROM quni_Memoranda GROUP BY dir ) AS grp " & _
                "ON grp.max_et = uni.sysEndTime AND grp.d = uni.dir " & _
                "ORDER BY uni.sysEndTime DESC, uni.sysStartTime DESC;"
End Property

Public Property Get Memoranda_byCriteria() As String
    Memoranda_byCriteria = "SELECT *, DMax(""sysStartTime"",""tbl__Comments"",""object='"" & [object] & ""'"") AS comment  FROM quni_Memoranda AS uni " & _
                "INNER JOIN ( SELECT dir AS d, MAX(sysEndTime) AS max_et FROM quni_Memoranda GROUP BY dir ) AS grp " & _
                "ON grp.max_et = uni.sysEndTime AND grp.d = uni.dir " & vbCrLf & _
                "WHERE FALSE {0}" & vbCrLf & _
                "ORDER BY uni.sysEndTime DESC, uni.sysStartTime DESC;"
End Property

Public Property Get Memoranda_quni() As String
    Memoranda_quni = "SELECT * FROM tbl_Memoranda_old UNION ALL SELECT * FROM tbl_Memoranda ORDER BY sysEndTime DESC , dir DESC;"
End Property

Public Property Get Organisations_current() As String
    Organisations_current = "SELECT *, DMax(""sysStartTime"",""tbl__Comments"",""object='"" & [object] & ""'"") AS comment  FROM tbl_Organisations ORDER BY sysStartTime DESC, dir DESC;"
End Property

Public Property Get Organisations_old() As String
    Organisations_old = "SELECT *, DMax(""sysStartTime"",""tbl__Comments"",""object='"" & [object] & ""'"") AS comment  FROM quni_Organisations AS uni " & _
                "INNER JOIN ( SELECT dir AS d, MAX(sysEndTime) AS max_et FROM quni_Organisations GROUP BY dir ) AS grp " & _
                "ON grp.max_et = uni.sysEndTime AND grp.d = uni.dir " & _
                "ORDER BY uni.sysEndTime DESC, uni.sysStartTime DESC;"
End Property

Public Property Get Organisations_byCriteria() As String
    Organisations_byCriteria = "SELECT *, DMax(""sysStartTime"",""tbl__Comments"",""object='"" & [object] & ""'"") AS comment  FROM quni_Organisations AS uni " & _
                "INNER JOIN ( SELECT dir AS d, MAX(sysEndTime) AS max_et FROM quni_Organisations GROUP BY dir ) AS grp " & _
                "ON grp.max_et = uni.sysEndTime AND grp.d = uni.dir " & vbCrLf & _
                "WHERE FALSE {0}" & vbCrLf & _
                "ORDER BY uni.sysEndTime DESC, uni.sysStartTime DESC;"
End Property

Public Property Get Organisations_quni() As String
    Organisations_quni = "SELECT * FROM tbl_Organisations_old UNION ALL SELECT * FROM tbl_Organisations ORDER BY sysEndTime DESC , dir DESC;"
End Property

Public Property Get Outputs_current() As String
    Outputs_current = "SELECT *, DMax(""sysStartTime"",""tbl__Comments"",""object='"" & [object] & ""'"") AS comment  FROM tbl_Outputs ORDER BY sysStartTime DESC, dir DESC;"
End Property

Public Property Get Outputs_old() As String
    Outputs_old = "SELECT *, DMax(""sysStartTime"",""tbl__Comments"",""object='"" & [object] & ""'"") AS comment  FROM quni_Outputs AS uni " & _
                "INNER JOIN ( SELECT dir AS d, MAX(sysEndTime) AS max_et FROM quni_Outputs GROUP BY dir ) AS grp " & _
                "ON grp.max_et = uni.sysEndTime AND grp.d = uni.dir " & _
                "ORDER BY uni.sysEndTime DESC, uni.sysStartTime DESC;"
End Property

Public Property Get Outputs_byCriteria() As String
    Outputs_byCriteria = "SELECT *, DMax(""sysStartTime"",""tbl__Comments"",""object='"" & [object] & ""'"") AS comment  FROM quni_Outputs AS uni " & _
                "INNER JOIN ( SELECT dir AS d, MAX(sysEndTime) AS max_et FROM quni_Outputs GROUP BY dir ) AS grp " & _
                "ON grp.max_et = uni.sysEndTime AND grp.d = uni.dir " & vbCrLf & _
                "WHERE FALSE {0}" & vbCrLf & _
                "ORDER BY uni.sysEndTime DESC, uni.sysStartTime DESC;"
End Property

Public Property Get Outputs_quni() As String
    Outputs_quni = "SELECT * FROM tbl_Outputs_old UNION ALL SELECT * FROM tbl_Outputs ORDER BY sysEndTime DESC , dir DESC;"
End Property

Public Property Get People_current() As String
    People_current = "SELECT *, DMax(""sysStartTime"",""tbl__Comments"",""object='"" & [object] & ""'"") AS comment  FROM tbl_People ORDER BY sysStartTime DESC, dir DESC;"
End Property

Public Property Get People_old() As String
    People_old = "SELECT *, DMax(""sysStartTime"",""tbl__Comments"",""object='"" & [object] & ""'"") AS comment  FROM quni_People AS uni " & _
                "INNER JOIN ( SELECT dir AS d, MAX(sysEndTime) AS max_et FROM quni_People GROUP BY dir ) AS grp " & _
                "ON grp.max_et = uni.sysEndTime AND grp.d = uni.dir " & _
                "ORDER BY uni.sysEndTime DESC, uni.sysStartTime DESC;"
End Property

Public Property Get People_byCriteria() As String
    People_byCriteria = "SELECT *, DMax(""sysStartTime"",""tbl__Comments"",""object='"" & [object] & ""'"") AS comment  FROM quni_People AS uni " & _
                "INNER JOIN ( SELECT dir AS d, MAX(sysEndTime) AS max_et FROM quni_People GROUP BY dir ) AS grp " & _
                "ON grp.max_et = uni.sysEndTime AND grp.d = uni.dir " & vbCrLf & _
                "WHERE FALSE {0}" & vbCrLf & _
                "ORDER BY uni.sysEndTime DESC, uni.sysStartTime DESC;"
End Property

Public Property Get People_quni() As String
    People_quni = "SELECT * FROM tbl_People_old UNION ALL SELECT * FROM tbl_People ORDER BY sysEndTime DESC , dir DESC;"
End Property

Public Property Get Schemes_current() As String
    Schemes_current = "SELECT *, DMax(""sysStartTime"",""tbl__Comments"",""object='"" & [object] & ""'"") AS comment FROM tbl_Schemes WHERE schemeCode = Forms![frm_Schemes]![cmb_Main] ORDER BY sort;"
End Property

Public Property Get Schemes_old() As String
    Schemes_old = "SELECT *, DMax(""sysStartTime"",""tbl__Comments"",""object='"" & [object] & ""'"") AS comment  FROM quni_Schemes AS uni " & _
          "INNER JOIN ( SELECT dir AS d, MAX(sysEndTime) AS max_et FROM quni_Schemes GROUP BY dir ) AS grp " & _
          "ON grp.max_et = uni.sysEndTime AND grp.d = uni.dir " & _
          "WHERE schemeCode = Forms![frm_Schemes]![cmb_Main] " & _
          "ORDER BY uni.sysEndTime DESC, sort;"
End Property

Public Property Get Schemes_quni() As String
    Schemes_quni = "SELECT * FROM tbl_Schemes_old UNION ALL SELECT * FROM tbl_Schemes ORDER BY sysEndTime DESC , dir DESC;"
End Property

Public Property Get WordPad_quni() As String
    WordPad_quni = "SELECT * FROM tbl_WordPad_old UNION ALL SELECT * FROM tbl_WordPad ORDER BY sysEndTime DESC , object DESC;"
End Property

Public Function Where_AttachedToAlike(csv As String) As String
    Dim str() As String
    Dim i As Long: i = 0
    Dim s As Variant
    For Each s In Split(csv, ",")
        ReDim Preserve str(i)
        str(i) = Replace("OR attachedTo Alike '%,{0},%'", "{0}", s)
        i = i + 1
    Next s
    Where_AttachedToAlike = Join(str, vbCrLf)
End Function

Public Function Where_SchemeEquals(csv As String) As String
    Dim str() As String
    Dim i As Long: i = 0
    Dim s As Variant
    For Each s In Split(csv, ",")
        ReDim Preserve str(i)
        str(i) = Replace("OR scheme = {0}", "{0}", s)
        i = i + 1
    Next s
    Where_SchemeEquals = Join(str, vbCrLf)
End Function

Public Property Get Comments_1() As String
    Comments_1 = "SELECT * FROM tbl__Comments WHERE object = '{0}' ORDER BY sysStartTime DESC;"
End Property

'Public Property Get Comments_2() As String
'    Comments_2 = "TRANSFORM Count(tbl__Comments.ID) AS CountOfID " & _
'                 "SELECT Int([sysStartTime]) AS [Day] FROM tbl__Comments " & _
'                 "WHERE (((tbl__Comments.object)='{0}') AND (((tbl__Comments.[comment]) = '_clickDir' Or (tbl__Comments.[comment]) = '_clickLink'))) " & _
'                 "GROUP BY Int([sysStartTime]) " & _
'                 "ORDER BY Int([sysStartTime]) DESC " & _
'                 "PIVOT tbl__Comments.[comment];"
'End Property

'Public Property Get Comments_3() As String
'    Comments_3 = "TRANSFORM Count(tbl__Comments.ID) AS CountOfID " & _
'                 "SELECT Int([sysStartTime]) AS [Day] FROM tbl__Comments " & _
'                 "WHERE (((tbl__Comments.object)='{0}') AND NOT (((tbl__Comments.[comment]) = '_clickDir' Or (tbl__Comments.[comment]) = '_clickLink'))) " & _
'                 "GROUP BY Int([sysStartTime]) " & _
'                 "ORDER BY Int([sysStartTime]) DESC " & _
'                 "PIVOT tbl__Comments.[comment];"
'End Property
