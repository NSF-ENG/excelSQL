VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Reviewers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit
Private QT As clsQueryTable

Function Descriptor() As clsSheetDescriptor
' return descriptor for this sheet
    Set Descriptor = New clsSheetDescriptor
    With Descriptor
     .name = Me.name
     .description = "[slow]Reviewers who have completed reviews"
     .tip = "Slow if retrieving statistics"
     .tabtitle = "Reviewers who have done reviews"
     .note = "Reviewers who have completed reviews on at least one of the chosen projects.  Can be used to check reviews on a panel, or to find prospective panelists."
     .helpText = "Shows reviewers of the chosen proposals, along with the number done and some statistics.  Reviewers with blank institutions should probably be updated.  This is intended for creating mailing lists, but can also be used to, say, track the number of reviews recieved for a panel.  As NSF adds more links between the PI and Reviewer databases, this will compare the emails stored."
     .helpvideo = ""
     .phase = 3
     .order = 3103
     .uigroup = .order Mod 10
     .depends = InputTab.name
    End With
End Function

Sub InitializeQueryTable()
' record query table for use by QT.doRefresh and events BeforeRefresh and AfterRefresh
' Called in ThisWorkbook_Open with on Error resume next so turn that off first.
  On Error GoTo 0
  Set QT = New clsQueryTable
  Call QT.initialize(Me, Me.ListObjects(1).QueryTable)
End Sub

Sub ClearRS()
  If Not QT Is Nothing Then Call QT.clear
End Sub

Sub RefreshRS(Optional limit As String = "")
'refresh tab based on parameters in InputTab
Dim dropTables As String
Dim mainQuery As String
Dim revtable As String

revtable = "create table #revScores(yn char(5) primary key, string varchar(10), score real null)" & vbNewLine & "insert into #revScores" & vbNewLine _
& "select 'NNNNN', 'R', null  union all" & vbNewLine & "select 'NNNNY', 'P', 1 union all" & vbNewLine & "select 'NNNYN', 'F', 3 union all" & vbNewLine _
& "select 'NNNYY', 'F/P', 2 union all" & vbNewLine & "select 'NNYNN', 'G', 5 union all" & vbNewLine & "select 'NNYNY', 'G/P', 2.98 union all" & vbNewLine _
& "select 'NNYYN', 'G/F', 4 union all" & vbNewLine & "select 'NNYYY', 'G/F/P', 2.99 union all" & vbNewLine & "select 'NYNNN', 'V', 7 union all" & vbNewLine _
& "select 'NYNNY', 'V/P', 3.98 union all" & vbNewLine & "select 'NYNYN', 'V/F', 4.98 union all" & vbNewLine & "select 'NYNYY', 'V/F/P', 3.65 union all" & vbNewLine _
& "select 'NYYNN', 'V/G', 6 union all" & vbNewLine & "select 'NYYNY', 'V/G/P', 4.32 union all" & vbNewLine & "select 'NYYYN', 'V/G/F', 4.99 union all" & vbNewLine _
& "select 'NYYYY', 'V/G/F/P', 3.97 union all" & vbNewLine & "select 'YNNNN', 'E', 9 union all" & vbNewLine & "select 'YNNNY', 'E/P', 4.992 union all" & vbNewLine _
& "select 'YNNYN', 'E/F', 5.98 union all" & vbNewLine & "select 'YNNYY', 'E/F/P', 4.325 union all" & vbNewLine & "select 'YNYNN', 'E/G', 6.98 union all" & vbNewLine _
& "select 'YNYNY', 'E/G/P', 4.995 union all" & vbNewLine & "select 'YNYYN', 'E/G/F', 5.66 union all" & vbNewLine & "select 'YNYYY', 'E/G/F/P', 4.5 union all" & vbNewLine _
& "select 'YYNNN', 'E/V', 8 union all" & vbNewLine & "select 'YYNNY', 'E/V/P', 5.666 union all" & vbNewLine & "select 'YYNYN', 'E/V/F', 6.33 union all" & vbNewLine _
& "select 'YYNYY', 'E/V/F/P', 4.996 union all" & vbNewLine & "select 'YYYNN', 'E/V/G', 6.99 union all" & vbNewLine & "select 'YYYNY', 'E/V/G/P', 5.5 union all" & vbNewLine _
& "select 'YYYYN', 'E/V/G/F', 5.99 union all" & vbNewLine & "select 'YYYYY', 'E/V/G/F/P', 4.997" & vbNewLine

'reviewers on these proposals
If Len(Me.Range("reviews_since_date").Value) < 2 Then
    mainQuery = "SELECT  " & limit & " r.revr_id, COUNT(r.prop_id) as nProps, MAX(r.rev_rtrn_date) as latest, revr.revr_last_name, revr.revr_frst_name, revr.gend_code," & vbNewLine _
    & "revr.inst_id, inst.inst_shrt_name, inst.st_code, revr.revr_dept_name, f.fos_txt as field, " & vbNewLine _
    & "'' as 'min_score', '' as 'avg_score', '' as 'max_score', '' as 'std_score'," & vbNewLine _
    & "'' AS 'avg_len', '' AS 'std_len', '' AS 'avg_days_early', ''  AS 'std_days_early'," & vbNewLine _
    & "ra.revr_addr_txt as 'revr_email', revr.pi_id, pi.pi_emai_addr" & vbNewLine _
    & "FROM (SELECT rp.revr_id, rp.prop_id, rp.rev_rtrn_date " & vbNewLine _
    & "    FROM (SELECT DISTINCT lead FROM #myProps) p" & vbNewLine _
    & "    JOIN csd.rev_prop rp ON rp.prop_id = p.lead" & vbNewLine _
    & "    WHERE rp.rev_rtrn_date is not NULL and rp.rev_type_code <> 'C') r" & vbNewLine _
    & "JOIN csd.revr revr ON revr.revr_id = r.revr_id" & vbNewLine _
    & "LEFT OUTER JOIN csd.fos f on f.prmy_fos_code = revr.prmy_fos_code" & vbNewLine _
    & "LEFT OUTER JOIN csd.revr_opt_addr_line ra ON ra.revr_id = r.revr_id AND ra.addr_lne_type_code='E'" & vbNewLine _
    & "LEFT OUTER JOIN csd.inst inst ON revr.inst_id = inst.inst_id " & vbNewLine _
    & "LEFT OUTER JOIN csd.pi_vw pi ON revr.pi_id = pi.pi_id" & vbNewLine _
    & "GROUP BY r.revr_id, revr.revr_last_name, revr.revr_frst_name, revr.gend_code, f.fos_txt, revr.inst_id, inst.inst_shrt_name, inst.st_code, revr.revr_dept_name, ra.revr_addr_txt, revr.pi_id, pi.pi_emai_addr" & vbNewLine _
    & "ORDER BY nProps DESC, revr.revr_last_name, revr.revr_frst_name" & vbNewLine

    dropTables = "DROP TABLE #myProps DROP TABLE #revScores"
Else
    mainQuery = revtable & "SELECT  " & limit & " r.revr_id, COUNT(r.prop_id) as nProps, MAX(r.rev_rtrn_date) as latest, revr.revr_last_name, revr.revr_frst_name, revr.gend_code," & vbNewLine _
    & "revr.inst_id, inst.inst_shrt_name, inst.st_code, revr.revr_dept_name, f.fos_txt as field, " & vbNewLine _
    & "min(r.score) as 'min_score', avg(r.score) as 'avg_score', max(r.score) as 'max_score', stdev(r.score) as 'std_score'," & vbNewLine _
    & "avg(r.length) AS 'avg_len', stdev(r.length) AS 'std_len', avg(r.days_early) AS 'avg_days_early', stdev(r.days_early)  AS 'std_days_early'," & vbNewLine _
    & "ra.revr_addr_txt as 'revr_email', revr.pi_id, pi.pi_emai_addr" & vbNewLine _
    & "FROM (SELECT rp.revr_id, rp.prop_id, rs.score, rp.rev_rtrn_date," & vbNewLine _
    & "  len(rp_txt_flds_vw.REV_PROP_TXT_FLDS) AS 'length', convert(float,DATEDIFF(day, rp.rev_rtrn_date, rp.rev_due_date)) AS 'days_early'" & vbNewLine _
    & "    FROM (SELECT DISTINCT lead FROM #myProps) p" & vbNewLine _
    & "    JOIN csd.rev_prop rp ON rp.prop_id = p.lead" & vbNewLine _
    & "    JOIN csd.rev_prop_txt_flds_vw rp_txt_flds_vw ON rp_txt_flds_vw.PROP_ID = p.lead AND rp_txt_flds_vw.REVR_ID = rp.revr_id" & vbNewLine _
    & "    JOIN csd.rev_prop_vw rp_vw ON rp_vw.PROP_ID = p.lead AND rp_vw.REVR_ID = rp.revr_id" & vbNewLine _
    & "    JOIN #revScores rs ON rp_vw.rev_prop_rtng_ind = rs.yn" & vbNewLine _
    & "    WHERE rp.rev_rtrn_date is not NULL and rp.rev_type_code <> 'C') r" & vbNewLine _
    & "JOIN csd.revr revr ON revr.revr_id = r.revr_id" & vbNewLine _
    & "LEFT OUTER JOIN csd.fos f on f.prmy_fos_code = revr.prmy_fos_code" & vbNewLine _
    & "LEFT OUTER JOIN csd.revr_opt_addr_line ra ON ra.revr_id = r.revr_id AND ra.addr_lne_type_code='E'" & vbNewLine _
    & "LEFT OUTER JOIN csd.inst inst ON revr.inst_id = inst.inst_id " & vbNewLine _
    & "LEFT OUTER JOIN csd.pi_vw pi ON revr.pi_id = pi.pi_id" & vbNewLine _
    & "GROUP BY r.revr_id, revr.revr_last_name, revr.revr_frst_name, revr.gend_code, f.fos_txt, revr.inst_id, inst.inst_shrt_name, inst.st_code, revr.revr_dept_name, ra.revr_addr_txt, revr.pi_id, pi.pi_emai_addr" & vbNewLine _
    & "ORDER BY nProps DESC, revr.revr_last_name, revr.revr_frst_name" & vbNewLine
    
    dropTables = "DROP TABLE #myProps DROP TABLE #revScores"
End If

Call QT.doQuery(mInputTab.InputSQL() & mainQuery & dropTables)
End Sub

