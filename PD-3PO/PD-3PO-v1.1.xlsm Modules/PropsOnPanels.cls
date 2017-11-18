VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PropsOnPanels"
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
     .description = "Panel schedule and results"
     .tip = "Check prop_stts"
     .tabtitle = "Proposals on Panels"
     .note = "Gets panel schedule and results for the chosen proposals; buttons sort by schedule or results.  Proposals on several panels are repeated."
     .helpText = "Shows panels handling the chosen proposals. Can sort by dates & discussion order to know when you want to attend, or by recommendation and score to see results. Note: All review scores are listed with each panel to make sure ad hocs are considered."
     .helpvideo = ""
     .phase = 3
     .order = 424
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
Dim myRevs, myPanls, Query As String
Dim revtable As String


'Props+panel results
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

myRevs = "SELECT p.lead, rs.string, rs.score, rev_prop_vw.revr_id, rev_prop_vw.rev_rtrn_date, id=identity(18), 0 as 'seq'" & vbNewLine _
& "INTO #myRevs" & vbNewLine _
& "FROM #myProps p" & vbNewLine _
& "JOIN csd.rev_prop_vw rev_prop_vw ON p.lead = rev_prop_vw.prop_id" & vbNewLine _
& "JOIN csd.rev_prop rev_prop ON p.lead = rev_prop.prop_id AND rev_prop_vw.revr_id = rev_prop.revr_id" & vbNewLine _
& "JOIN #revScores rs ON rev_prop_vw.rev_prop_rtng_ind = rs.yn " & vbNewLine _
& "WHERE p.ILN < 'M' AND rev_prop.rev_stts_code <> 'C'" & vbNewLine _
& "ORDER BY lead, score DESC" & vbNewLine _
& "SELECT lead, MIN(id) as 'start' INTO #myStarts FROM #myRevs GROUP BY lead" & vbNewLine _
& "UPDATE #myRevs set seq = id-M.start FROM #myRevs r, #myStarts M WHERE r.lead = M.lead" & vbNewLine _
& "DROP TABLE #myStarts" & vbNewLine

myPanls = "SELECT prop.lead, pp.prop_seq_num as disc_ordr, pp.panl_id, CASE WHEN pps.panl_id IS NOT NULL THEN 'Y' END AS panl_held, " & vbNewLine _
& "pps.panl_summ_rlse_flag as summ_rlse, pps.PROP_ORDR as rank, pps.RCOM_SEQ_NUM as rcom_seq, prd.rcom_abbr, prd.rcom_txt" & vbNewLine _
& "INTO #myPanlProp " & vbNewLine _
& "FROM #myProps prop" & vbNewLine _
& "JOIN csd.panl_prop pp ON prop.lead = pp.prop_id" & vbNewLine _
& "LEFT OUTER JOIN FLflpdb.flp.panl_prop_summ pps ON prop.lead = pps.PROP_ID AND pp.panl_id = pps.panl_id" & vbNewLine _
& "LEFT OUTER JOIN FLflpdb.flp.panl_rcom_def prd ON pp.panl_id = prd.panl_id AND pps.RCOM_SEQ_NUM = prd.RCOM_SEQ_NUM" & vbNewLine _
& "WHERE prop.ILN < 'M' " & vbNewLine

myPanls = myPanls & "SELECT pl.panl_id, panl.panl_name, panl.panl_bgn_date, panl.panl_loc, convert(varchar,SUM(ps.rtCount)) + ' proj: ' +" & vbNewLine _
& "   MAX( CASE ps.RCOM_SEQ_NUM WHEN 1 THEN        convert(varchar,ps.rtCount) + ' ' +  ps.RCOM_ABBR ELSE '' END ) +" & vbNewLine _
& "   MAX( CASE ps.RCOM_SEQ_NUM WHEN 2 THEN ', ' + convert(varchar,ps.rtCount) + ' ' +  ps.RCOM_ABBR ELSE '' END ) +" & vbNewLine _
& "   MAX( CASE ps.RCOM_SEQ_NUM WHEN 3 THEN ', ' + convert(varchar,ps.rtCount) + ' ' +  ps.RCOM_ABBR ELSE '' END ) +" & vbNewLine _
& "   MAX( CASE ps.RCOM_SEQ_NUM WHEN 4 THEN ', ' + convert(varchar,ps.rtCount) + ' ' +  ps.RCOM_ABBR ELSE '' END ) +" & vbNewLine _
& "   MAX( CASE ps.RCOM_SEQ_NUM WHEN 5 THEN ', ' + convert(varchar,ps.rtCount) + ' ' +  ps.RCOM_ABBR ELSE '' END ) +" & vbNewLine _
& "   MAX( CASE ps.RCOM_SEQ_NUM WHEN 6 THEN ', ' + convert(varchar,ps.rtCount) + ' ' +  ps.RCOM_ABBR ELSE '' END ) as 'panlSumm'" & vbNewLine _
& "INTO #myPanls" & vbNewLine _
& "FROM (SELECT DISTINCT panl_id FROM #myPanlProp) pl" & vbNewLine _
& "JOIN csd.panl panl ON pl.panl_id = panl.panl_id" & vbNewLine _
& "LEFT OUTER JOIN (SELECT p2.panl_id, panl_rcom_def.RCOM_SEQ_NUM, panl_rcom_def.RCOM_ABBR, Count(panl_prop_summ.PROP_ID) AS rtCount" & vbNewLine _
& "    FROM (SELECT DISTINCT panl_id FROM #myPanlProp) p2" & vbNewLine _
& "    JOIN FLflpdb.flp.panl_prop_summ panl_prop_summ ON p2.panl_id = panl_prop_summ.PANL_ID" & vbNewLine _
& "    JOIN FLflpdb.flp.panl_rcom_def panl_rcom_def ON panl_prop_summ.PANL_ID = panl_rcom_def.PANL_ID AND panl_prop_summ.RCOM_SEQ_NUM = panl_rcom_def.RCOM_SEQ_NUM" & vbNewLine _
& "    JOIN csd.prop pr ON panl_prop_summ.PROP_ID = pr.prop_id AND pr.prop_id=isnull(pr.lead_prop_id,pr.prop_id) " & vbNewLine _
& "    GROUP BY p2.panl_id, panl_rcom_def.RCOM_SEQ_NUM, panl_rcom_def.RCOM_ABBR ) ps ON pl.panl_id = ps.panl_id" & vbNewLine _
& "GROUP BY pl.panl_id, panl.panl_name, panl.panl_bgn_date, panl.panl_loc " & vbNewLine

Query = "SELECT " & limit & " prop.*, pp.panl_id, pn.panl_bgn_date, pn.panl_loc, pp.disc_ordr, pn.panl_name, " & vbNewLine _
& "pp.panl_held, pp.summ_rlse, pp.rank, pp.rcom_seq, pp.rcom_abbr, pp.rcom_txt," & vbNewLine _
& "CASE WHEN pp.rank IS NOT NULL THEN 'rk '+convert(varchar,pp.rank)+' of ' ELSE '' END + pn.panlSumm AS rankInPanl, " & vbNewLine _
& "CASE WHEN PROP_COVR.HUM_DATE is not NULL THEN convert(varchar(10),PROP_COVR.HUM_DATE,1) WHEN PROP_COVR.humn_date_pend_flag='Y' THEN 'Pend' END AS humn_date," & vbNewLine _
& "CASE WHEN PROP_COVR.VERT_DATE is not NULL THEN convert(varchar(10),PROP_COVR.VERT_DATE,1) WHEN PROP_COVR.vrtb_date_pend_flag='Y' THEN 'Pend' END AS vrtb_date," & vbNewLine _
& "revs.reviews, revs.avg_score, revs.last_rev_date," & vbNewLine

Query = Query & "(SELECT MAX(CASE b.seq WHEN 1 THEN b.ctry_name ELSE '' END)+" & vbNewLine _
& "    MAX(CASE b.seq WHEN 2 THEN '; '+b.ctry_name ELSE '' END)+" & vbNewLine _
& "    MAX(CASE b.seq WHEN 3 THEN '; '+b.ctry_name ELSE '' END)+" & vbNewLine _
& "    MAX(CASE b.seq WHEN 4 THEN '; '+b.ctry_name ELSE '' END)+" & vbNewLine _
& "    MAX(CASE b.seq WHEN 5 THEN '; '+b.ctry_name ELSE '' END)+" & vbNewLine _
& "    MAX(CASE b.seq WHEN 6 THEN '; '+b.ctry_name ELSE '' END)+" & vbNewLine _
& "    MAX(CASE b.seq WHEN 7 THEN '; '+b.ctry_name ELSE '' END)+" & vbNewLine _
& "    MAX(CASE b.seq WHEN 8 THEN '; '+b.ctry_name ELSE '' END) " & vbNewLine _
& "    FROM (SELECT prop_spcl_item_vw.TEMP_PROP_ID, prop_spcl_item_vw.SPCL_ITEM_CODE, ctry.ctry_name, " & vbNewLine _
& "            (SELECT COUNT(*) FROM csd.prop_spcl_item_vw sp2" & vbNewLine _
& "             JOIN csd.ctry ctry2 ON sp2.SPCL_ITEM_CODE = ctry2.ctry_code AND ctry.ctry_code >= ctry2.ctry_code" & vbNewLine _
& "             WHERE sp2.TEMP_PROP_ID = prop_spcl_item_vw.TEMP_PROP_ID) as seq" & vbNewLine _
& "          FROM csd.prop_spcl_item_vw prop_spcl_item_vw" & vbNewLine _
& "          JOIN csd.ctry ctry ON prop_spcl_item_vw.SPCL_ITEM_CODE = ctry.ctry_code ) b " & vbNewLine _
& "     WHERE b.TEMP_PROP_ID = prop_subm_ctl_vw.TEMP_PROP_ID" & vbNewLine _
& "     GROUP BY b.TEMP_PROP_ID) AS country " & vbNewLine

Query = Query & "FROM #myProps prop" & vbNewLine _
& "JOIN csd.prop_subm_ctl_vw prop_subm_ctl_vw ON prop.prop_id = prop_subm_ctl_vw.PROP_ID" & vbNewLine _
& "JOIN FLflpdb.flp.PROP_COVR PROP_COVR ON prop_subm_ctl_vw.TEMP_PROP_ID = PROP_COVR.TEMP_PROP_ID" & vbNewLine _
& "LEFT OUTER JOIN #myPanlProp pp ON prop.lead = pp.lead" & vbNewLine _
& "LEFT OUTER JOIN #myPanls pn ON pp.panl_id = pn.panl_id " & vbNewLine _
& "LEFT OUTER JOIN (SELECT r.lead, AVG(r.score) AS avg_score, MAX(r.rev_rtrn_date) AS last_rev_date, MAX(CASE r.seq WHEN 0 THEN r.string ELSE '' END)+" & vbNewLine _
& "      MAX(CASE r.seq WHEN 1 THEN ','+r.string ELSE '' END)+" & vbNewLine _
& "      MAX(CASE r.seq WHEN 2 THEN ','+r.string ELSE '' END)+" & vbNewLine _
& "      MAX(CASE r.seq WHEN 3 THEN ','+r.string ELSE '' END)+" & vbNewLine _
& "      MAX(CASE r.seq WHEN 4 THEN ','+r.string ELSE '' END)+" & vbNewLine _
& "      MAX(CASE r.seq WHEN 5 THEN ','+r.string ELSE '' END)+" & vbNewLine _
& "      MAX(CASE r.seq WHEN 6 THEN ','+r.string ELSE '' END)+" & vbNewLine _
& "      MAX(CASE r.seq WHEN 7 THEN ','+r.string ELSE '' END)+" & vbNewLine _
& "      MAX(CASE r.seq WHEN 8 THEN ','+r.string ELSE '' END)+" & vbNewLine _
& "      MAX(CASE r.seq WHEN 9 THEN ','+r.string ELSE '' END) AS reviews" & vbNewLine _
& "    FROM #myRevs r " & vbNewLine _
& "    GROUP BY r.lead) revs ON prop.lead = revs.lead" & vbNewLine _
& "ORDER BY pn.panl_bgn_date, pn.panl_id, pp.disc_ordr, prop.lead, prop.ILN" & vbNewLine

dropTables = "DROP TABLE #myProps DROP TABLE #myPanls DROP TABLE #revScores DROP TABLE #myRevs" & vbNewLine
Call QT.doQuery(mInputTab.InputSQL & revtable & myRevs & myPanls & Query & dropTables)
End Sub
