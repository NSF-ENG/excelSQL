Attribute VB_Name = "mQueries"
Option Explicit
Function pidHeader(template As String) As String
' common header for proposal ids.
pidHeader = "SELECT isnull(lead_prop_id,prop_id) AS lead," & vbNewLine _
& "CASE WHEN lead_prop_id IS NULL THEN 'I' WHEN lead_prop_id = prop_id THEN 'L' ELSE 'N' END AS ILN, prop.prop_id," & vbNewLine _
& "prop_id"
End Function

Function propHeader() As String ' common header, used in both refresh subs
propHeader = "SELECT CASE WHEN p.lead_prop_id IS NULL THEN 'I' WHEN p.lead_prop_id <> p.prop_id THEN 'N' ELSE 'L' END AS ILN," & vbNewLine _
& "isnull(p.lead_prop_id,p.prop_id) AS lead, p.prop_id, psc.TEMP_PROP_ID" & vbNewLine _
& "INTO #myPids FROM csd.prop p" & vbNewLine _
& "JOIN csd.prop_subm_ctl_vw psc ON psc.prop_id = p.prop_id" & vbNewLine
End Function


Function propTrailer() As String ' common header, used in both refresh subs

propTrailer = "ORDER BY lead, ILN, p.prop_id" & vbNewLine _
& "CREATE INDEX myPid_idx ON #myPids(prop_id)" & vbNewLine _
& "-- we should have everyone already, but let's make sure" & vbNewLine _
& "INSERT INTO #myPids SELECT CASE WHEN pr.lead_prop_id <> pr.prop_id THEN 'N' ELSE 'L' END as ILN, " & vbNewLine _
& "p.lead, p.prop_id, psc.TEMP_PROP_ID" & vbNewLine _
& "FROM #myPids p " & vbNewLine _
& "JOIN csd.prop pr ON pr.lead_prop_id = p.lead" & vbNewLine _
& "JOIN csd.prop_subm_ctl_vw psc ON psc.prop_id = pr.prop_id" & vbNewLine _
& "WHERE p.ILN <> 'I' " & vbNewLine _
& "AND NOT EXISTS (SELECT * FROM #myPids px WHERE px.prop_id = pr.prop_id) -- skip what we have already" & vbNewLine _
& "ORDER BY lead, ILN, p.prop_id" & vbNewLine _
& "CREATE INDEX myPid_idx3 ON #myPids(ILN, lead, prop_id) " & vbNewLine _
& "CREATE INDEX myPid_temp ON #myPids(TEMP_PROP_ID)" & vbNewLine
End Function


Sub makeQueries(myProps As String)
Dim setNC As String
Dim qt As QueryTable
Dim dropTables As String

Call handlePwd

setNC = "SET NOCOUNT ON " & vbNewLine

doQuery(
tempTables = "SELECT p.prop_id, ctry.ctry_name, id=identity(18), 0 as 'seq' INTO #myCtry" & vbNewLine _
& "FROM #myPids p" & vbNewLine _
& "JOIN csd.prop_spcl_item_vw sp1 ON sp1.TEMP_PROP_ID = p.TEMP_PROP_ID" & vbNewLine _
& "JOIN csd.ctry ctry ON sp1.SPCL_ITEM_CODE = ctry.ctry_code" & vbNewLine _
& "WHERE end_date Is Null" & vbNewLine _
& "ORDER BY p.prop_id, ctry.ctry_name" & vbNewLine _
& "CREATE INDEX myCtry_idx ON #myCtry(prop_id) " & vbNewLine _
& "SELECT prop_id, MIN(id) as 'start' INTO #myStCtry FROM #myCtry GROUP BY prop_id" & vbNewLine _
& "UPDATE #myCtry set seq = id-M.start FROM #myCtry r, #myStCtry M WHERE r.prop_id = M.prop_id" & vbNewLine _
& "DROP TABLE #myStCtry" & vbNewLine _
& "SELECT DISTINCT p.prop_id, pa.prop_atr_code, id=identity(18), 0 as 'seq' INTO #myPRCs" & vbNewLine _
& "FROM #myPids p" & vbNewLine _
& "JOIN csd.prop_atr pa ON pa.prop_id = p.prop_id AND pa.prop_atr_type_code = 'PRC'" & vbNewLine _
& "ORDER BY p.prop_id, pa.prop_atr_code" & vbNewLine _
& "CREATE INDEX myPRCs_idx ON #myPRCs(prop_id) " & vbNewLine _
& "SELECT prop_id, MIN(id) as 'start' INTO #myPRCStart FROM #myPRCs GROUP BY prop_id" & vbNewLine _
& "UPDATE #myPRCs set seq = id-M.start FROM #myPRCs r, #myPRCStart M WHERE r.prop_id = M.prop_id" & vbNewLine _
& "DROP TABLE #myPRCStart" & vbNewLine

tempTables = tempTables & "SELECT prop_id, " & vbNewLine _
& "SUM(CASE WHEN pi_gend_code = 'F'THEN 1 ELSE 0 END) AS NfmlPIs," & vbNewLine _
& "SUM(CASE WHEN pi_ethn_code = 'H'THEN 1 ELSE 0 END) AS NhispPIs," & vbNewLine _
& "SUM(CASE WHEN dmog_tbl_code = 'H'AND dmog_code <> 'N' THEN 1 ELSE 0 END) AS NhndcpPIs," & vbNewLine _
& "SUM(CASE WHEN dmog_tbl_code = 'R'AND dmog_code NOT IN ('U','W','B3') THEN 1 ELSE 0 END) AS NnonWhtAsnPIs" & vbNewLine _
& "INTO #myDmog" & vbNewLine _
& "FROM (SELECT p.prop_id, pi_id FROM #myPids p JOIN csd.prop prop ON prop.prop_id = p.prop_id" & vbNewLine _
& "      UNION ALL SELECT p.prop_id, pi_id FROM #myPids p JOIN csd.addl_pi_invl a ON a.prop_id = p.prop_id) PIs" & vbNewLine _
& "LEFT OUTER JOIN csd.pi_vw pi ON pi.pi_id = PIs.pi_id" & vbNewLine _
& "LEFT OUTER JOIN csd.PI_dmog d ON d.pi_id = PIs.pi_id" & vbNewLine _
& "GROUP BY prop_id" & vbNewLine _
& "ORDER BY prop_id" & vbNewLine _
& "CREATE INDEX myDmog_idx ON #myDmog(prop_id)" & vbNewLine

tempTables = tempTables & "-- revs" & vbNewLine _
& "SELECT p.prop_id, COUNT(*) as Nrev," & vbNewLine _
& "nullif(SUM(CASE WHEN rpv.rev_prop_unrl_flag = 'Y' THEN 1 ELSE 0 END),0) as NrevUnreleasable," & vbNewLine _
& "nullif(SUM(CASE WHEN rpv.rev_rlse_flag = 'Y' THEN 0 WHEN rpv.rev_prop_unrl_flag = 'Y'THEN 0 ELSE 1 END),0) as NrevUnmarked " & vbNewLine _
& "INTO #myRevs " & vbNewLine _
& "FROM #myPids p" & vbNewLine _
& "JOIN csd.rev_prop rp ON rp.prop_id = p.prop_id AND rp.rev_stts_code <> 'C'" & vbNewLine _
& "JOIN csd.rev_prop_vw rpv ON rpv.prop_id = p.prop_id AND rpv.revr_id = rp.revr_id " & vbNewLine _
& "GROUP BY p.prop_id" & vbNewLine _
& "ORDER BY p.prop_id" & vbNewLine _
& "-- summ" & vbNewLine _
& "SELECT p.prop_id, COUNT(pp.panl_id) as Npanl," & vbNewLine _
& "nullif(SUM(CASE WHEN pps.panl_summ_unrl_flag = 'Y' THEN 1 ELSE 0 END),0) as NpanlUnreleasable," & vbNewLine _
& "nullif(SUM(CASE WHEN pps.panl_summ_rlse_flag = 'Y' THEN 0 WHEN pps.panl_summ_unrl_flag = 'Y'THEN 0 ELSE 1 END),0) as NpanlUnmarked  " & vbNewLine _
& "INTO #myPanl" & vbNewLine _
& "FROM #myPids p" & vbNewLine _
& "JOIN csd.panl_prop pp ON pp.prop_id = p.prop_id" & vbNewLine _
& "LEFT OUTER JOIN FLflpdb.flp.panl_prop_summ pps ON pps.prop_id = p.prop_id AND pps.panl_id = pp.panl_id" & vbNewLine _
& "GROUP BY p.prop_id" & vbNewLine _
& "ORDER BY p.prop_id" & vbNewLine



mainQuery = "-- props: get codes to check if they match leads" & vbNewLine _
& "SELECT p.nsf_rcvd_date, nullif(p.dd_rcom_date,'1900-01-01') AS dd_rcom_date," & vbNewLine _
& "ILN, lead, p.prop_id," & vbNewLine _
& "pi.pi_last_name, pi.pi_frst_name, inst.inst_shrt_name AS inst_name, " & vbNewLine _
& "pi.pi_emai_addr,p.rqst_dol, p.rqst_eff_date,p.rqst_mnth_cnt,p.cntx_stmt_id, p.bas_rsch_pct, p.apld_rsch_pct+p.educ_trng_pct+land_buld_fix_equp_pct+mjr_equp_pct+non_invt_pct AS other_pct," & vbNewLine _
& "CASE WHEN PC.HUM_DATE is not NULL THEN convert(varchar(10),PC.HUM_DATE,1) WHEN PC.humn_date_pend_flag='Y' THEN 'Pend' END AS humn_date," & vbNewLine _
& "CASE WHEN PC.VERT_DATE is not NULL THEN convert(varchar(10),PC.VERT_DATE,1) WHEN PC.vrtb_date_pend_flag='Y' THEN 'Pend' END AS vrtb_date," & vbNewLine _
& "(SELECT MAX(CASE b.seq WHEN 1 THEN b.ctry_name ELSE '' END)+" & vbNewLine _
& "    MAX(CASE b.seq WHEN 2 THEN '; '+b.ctry_name ELSE '' END)+" & vbNewLine _
& "    MAX(CASE b.seq WHEN 3 THEN '; '+b.ctry_name ELSE '' END)+" & vbNewLine _
& "    MAX(CASE b.seq WHEN 4 THEN '; '+b.ctry_name ELSE '' END)+" & vbNewLine _
& "    MAX(CASE b.seq WHEN 5 THEN '; '+b.ctry_name ELSE '' END)+" & vbNewLine _
& "    MAX(CASE b.seq WHEN 6 THEN '; '+b.ctry_name ELSE '' END)+" & vbNewLine _
& "    MAX(CASE b.seq WHEN 7 THEN '; '+b.ctry_name ELSE '' END)+" & vbNewLine _
& "    MAX(CASE b.seq WHEN 8 THEN '; '+b.ctry_name ELSE '' END) " & vbNewLine _
& "    FROM #myCtry b WHERE b.prop_id = prop.prop_id) AS Country," & vbNewLine _
& "a.abst_narr_txt, (SELECT SUM(frgn_trav_dol) FROM csd.eps_blip eb WHERE eb.prop_id = prop.prop_id AND NOT EXISTS (SELECT eb1.revn_num FROM csd.eps_blip eb1 WHERE eb.prop_id = eb1.prop_id AND eb.revn_num < eb1.revn_num)) as frgn_trvl_dol," & vbNewLine _
& "p.pgm_annc_id, p.org_code, p.pgm_ele_code, p.pm_ibm_logn_id as PO, NfmlPIs,NhispPIs,NhndcpPIs,NnonWhtAsnPIs, Nrev,NrevUnreleasable, NrevUnmarked,  Npanl, NpanlUnreleasable, NpanlUnmarked," & vbNewLine

mainQuery = mainQuery & "p.obj_clas_code, p.natr_rqst_code, p.prop_stts_code, PRCs, p.prop_titl_txt," & vbNewLine _
& "p.inst_id, inst.st_code, p.pi_id, ra.last_updt_tmsp as RAupdate" & vbNewLine _
& "FROM #myPids prop" & vbNewLine _
& "JOIN csd.prop p ON p.prop_id = prop.prop_id" & vbNewLine _
& "JOIN #myDmog d ON d.prop_id = prop.prop_id" & vbNewLine _
& "LEFT OUTER JOIN csd.inst inst ON inst.inst_id = p.inst_id" & vbNewLine _
& "LEFT OUTER JOIN csd.pi_vw pi ON pi.pi_id = p.pi_id" & vbNewLine _
& "LEFT OUTER JOIN #myRevs rv ON rv.prop_id = prop.prop_id" & vbNewLine _
& "LEFT OUTER JOIN #myPanl pl ON pl.prop_id = prop.prop_id" & vbNewLine _
& "LEFT OUTER JOIN csd.abst a ON a.awd_id = prop.prop_id" & vbNewLine _
& "LEFT OUTER JOIN csd.prop_rev_anly_vw ra ON ra.prop_id = prop.prop_id" & vbNewLine _
& "LEFT OUTER JOIN FLflpdb.flp.PROP_COVR PC ON PC.TEMP_PROP_ID = prop.TEMP_PROP_ID" & vbNewLine _
& "LEFT OUTER JOIN (SELECT prop_id," & vbNewLine _
& "        MAX( CASE pa.seq WHEN 0 THEN       pa.prop_atr_code ELSE '' END ) +" & vbNewLine _
& "        MAX( CASE pa.seq WHEN 1 THEN ' ' + pa.prop_atr_code ELSE '' END ) +" & vbNewLine _
& "        MAX( CASE pa.seq WHEN 2 THEN ' ' + pa.prop_atr_code ELSE '' END ) +" & vbNewLine _
& "        MAX( CASE pa.seq WHEN 3 THEN ' ' + pa.prop_atr_code ELSE '' END ) +" & vbNewLine _
& "        MAX( CASE pa.seq WHEN 4 THEN ' ' + pa.prop_atr_code ELSE '' END ) +" & vbNewLine _
& "        MAX( CASE pa.seq WHEN 5 THEN ' ' + pa.prop_atr_code ELSE '' END ) +" & vbNewLine _
& "        MAX( CASE pa.seq WHEN 6 THEN ' ' + pa.prop_atr_code ELSE '' END ) AS PRCs" & vbNewLine _
& "    FROM #myPRCs pa " & vbNewLine _
& "    GROUP BY prop_id) myPRCs ON myPRCs.prop_id = prop.prop_id" & vbNewLine _
& "WHERE pi.prim_addr_flag='Y'" & vbNewLine _
& "ORDER BY prop.lead, prop.ILN, prop.prop_id" & vbNewLine

dropTables = "DROP TABLE #myDmog DROP TABLE #myPids DROP TABLE #myPRCs DROP TABLE #myCtry DROP TABLE #myRevs DROP TABLE #myPanl" & vbNewLine

    Set qt = Worksheets("CheckCoding").ListObjects.Item(1).QueryTable
    With qt
        .CommandText = setNC & myProps & tempTables & mainQuery & dropTables
        'MsgBox Mid(.CommandText, 1000, 1000)
        .Refresh
    End With

tempTables = "SELECT p.prop_id, LTRIM(MAX( CASE pa.prop_atr_seq WHEN 0 THEN ' ' + pa.prop_atr_code ELSE '' END ) +" & vbNewLine _
& " MAX( CASE pa.prop_atr_seq WHEN 1 THEN ' ' + pa.prop_atr_code ELSE '' END ) +" & vbNewLine _
& " MAX( CASE pa.prop_atr_seq WHEN 2 THEN ' ' + pa.prop_atr_code ELSE '' END ) +" & vbNewLine _
& " MAX( CASE pa.prop_atr_seq WHEN 3 THEN ' ' + pa.prop_atr_code ELSE '' END ) + " & vbNewLine _
& " MAX( CASE pa.prop_atr_seq WHEN 4 THEN ' ' + pa.prop_atr_code ELSE '' END ) + " & vbNewLine _
& " MAX( CASE pa.prop_atr_seq WHEN 5 THEN ' ' + pa.prop_atr_code ELSE '' END ) + " & vbNewLine _
& " MAX( CASE pa.prop_atr_seq WHEN 6 THEN ' ' + pa.prop_atr_code ELSE '' END ) + " & vbNewLine _
& " MAX( CASE pa.prop_atr_seq WHEN 7 THEN ' ' + pa.prop_atr_code ELSE '' END ) + " & vbNewLine _
& " MAX( CASE pa.prop_atr_seq WHEN 8 THEN ' ' + pa.prop_atr_code ELSE '' END ) + " & vbNewLine _
& " MAX( CASE pa.prop_atr_seq WHEN 9 THEN ' ' + pa.prop_atr_code ELSE '' END )) AS PRCs" & vbNewLine _
& "INTO #myPRCs" & vbNewLine _
& "FROM #myPids p, csd.prop_atr pa WHERE p.prop_id = pa.prop_id AND pa.prop_atr_type_code = 'PRC'" & vbNewLine _
& "GROUP BY p.prop_id" & vbNewLine _
& "CREATE INDEX myPRCs_idx ON #myPRCs(prop_id)" & vbNewLine _
& "SELECT p.lead, p.prop_id, eb.revn_num, eb.budg_seq_yr, eb.budg_tot_dol, " & vbNewLine _
& "eb.sr_pers_cnt, eb.sr_summ_mnth_cnt, eb.pdoc_grnt_dol, eb.frgn_trav_dol, eb.part_dol" & vbNewLine _
& "INTO #myBudg" & vbNewLine _
& "FROM #myPids p, csd.eps_blip eb" & vbNewLine _
& "WHERE p.prop_id = eb.prop_id AND NOT EXISTS (SELECT eb1.revn_num FROM csd.eps_blip eb1 WHERE eb.prop_id = eb1.prop_id AND eb.revn_num < eb1.revn_num)" & vbNewLine _
& "ORDER BY prop_id,budg_seq_yr" & vbNewLine _
& "CREATE INDEX myBudg_idx ON #myBudg(prop_id, budg_seq_yr)" & vbNewLine

mainQuery = "SELECT p.nsf_rcvd_date, p.pgm_annc_id, p.org_code, p.pgm_ele_code, p.pm_ibm_logn_id, prop.lead,  prop.ILN, prop.prop_id, '' as pi_last_name, '' as inst_name, '' as pi_emai_addr, " & vbNewLine _
& "eb.sr_pers_cnt, eb.sr_summ_mnth_cnt, eb.pdoc_grnt_dol, eb.frgn_trav_dol, eb.part_dol, eb.budg_tot_dol, " & vbNewLine _
& "eb.revn_num, eb.budg_seq_yr, eb.budg_tot_dol AS dol, p.org_code, p.pgm_ele_code as PEC, prc.PRCs" & vbNewLine _
& "FROM #myPids prop" & vbNewLine _
& "JOIN csd.prop p ON p.prop_id = prop.prop_id" & vbNewLine _
& "LEFT OUTER JOIN #myPRCs prc ON prc.prop_id = prop.prop_id" & vbNewLine _
& "LEFT OUTER JOIN #myBudg eb ON eb.prop_id = prop.prop_id " & vbNewLine _
& "UNION ALL SELECT p.nsf_rcvd_date, p.pgm_annc_id, p.org_code, p.pgm_ele_code, p.pm_ibm_logn_id, prop.lead,  prop.ILN, prop.prop_id, pi.pi_last_name, inst.inst_shrt_name, pi.pi_emai_addr, " & vbNewLine _
& "null, null, null, null, null, null, (SELECT MAX(revn_num) from #myBudg b WHERE b.prop_id = prop.prop_id)," & vbNewLine _
& "null, (SELECT SUM(budg_tot_dol) from #myBudg b WHERE b.prop_id = prop.prop_id), p.org_code, p.pgm_ele_code, prc.PRCs " & vbNewLine _
& "FROM #myPids prop" & vbNewLine _
& "JOIN csd.prop p ON p.prop_id = prop.prop_id" & vbNewLine _
& "LEFT OUTER JOIN csd.pi_vw pi ON pi.pi_id = p.pi_id AND pi.prim_addr_flag='Y'" & vbNewLine _
& "LEFT OUTER JOIN csd.inst inst ON inst.inst_id = p.inst_id" & vbNewLine _
& "LEFT OUTER JOIN #myPRCs prc ON prc.prop_id = prop.prop_id" & vbNewLine _
& "ORDER BY prop.lead, prop.ILN, prop.prop_id, eb.budg_seq_yr " & vbNewLine

dropTables = "DROP TABLE #myPids" & vbNewLine & "DROP TABLE #myPRCs" & vbNewLine & "DROP TABLE #myBudg" & vbNewLine

    Set qt = Worksheets("BudgetBlocks").ListObjects.Item(1).QueryTable
    With qt
        .CommandText = setNC & myProps & tempTables & mainQuery & dropTables
        'MsgBox Mid(.CommandText, 1000, 1000)
        .Refresh
    End With
    

tempTables = "SELECT lead, convert(text, convert(varchar(16384),js.PROJ_SUMM_TXT) + convert(varchar(16384),js.INTUL_MERT) + convert(varchar(16384),js.BRODR_IMPT)) as summ" & vbNewLine _
& "INTO #mySumm" & vbNewLine _
& "FROM #myPids p " & vbNewLine _
& "JOIN FLflpdb.flp.proj_summ js ON js.SPCL_CHAR_PDF <> 'Y' AND js.TEMP_PROP_ID = p.TEMP_PROP_ID" & vbNewLine _
& "WHERE p.ILN < 'M'" & vbNewLine _
& "SELECT prop.nsf_rcvd_date, nullif(prop.dd_rcom_date,'1900-01-01') AS dd_rcom_date," & vbNewLine _
& "prop.pgm_annc_id, o2.dir_div_abbr as Dir, prop.org_code, CASE WHEN prop.org_code <> prop.orig_org_code THEN prop.orig_org_code END AS origORG, " & vbNewLine _
& "prop.pgm_ele_code+' - '+pgm_ele_name as Pgm, CASE WHEN prop.pgm_ele_code <> prop.orig_pgm_ele_code THEN prop.orig_pgm_ele_code ELSE ' ' END AS origPEC," & vbNewLine _
& "prop.pm_ibm_logn_id as PO, prop.obj_clas_code, natr_rqst.natr_rqst_abbr, prop_stts.prop_stts_abbr, p.ILN, p.lead, org.dir_div_abbr as Div, p.prop_id,p.TEMP_PROP_ID," & vbNewLine _
& "pi.pi_last_name, pi.pi_frst_name, inst.inst_shrt_name AS inst_name, inst.st_code, pi.pi_emai_addr," & vbNewLine _
& "prop.prop_titl_txt, prop.rqst_dol, prop.rqst_eff_date, prop.rqst_mnth_cnt, prop.cntx_stmt_id, prop.prop_stts_code, prop_stts.prop_stts_txt, prop.inst_id, prop.pi_id" & vbNewLine _
& "INTO #myProps FROM (SELECT DISTINCT * FROM #myPids ) p" & vbNewLine _
& "JOIN csd.prop prop ON p.prop_id = prop.prop_id" & vbNewLine _
& "JOIN csd.org org ON prop.org_code = org.org_code" & vbNewLine _
& "JOIN csd.org o2 ON left(prop.org_code,2)+'000000' = o2.org_code" & vbNewLine _
& "JOIN csd.pgm_ele pe ON prop.pgm_ele_code = pe.pgm_ele_code" & vbNewLine _
& "JOIN csd.inst inst ON prop.inst_id = inst.inst_id" & vbNewLine _
& "JOIN csd.pi_vw pi ON prop.pi_id = pi.pi_id" & vbNewLine _
& "JOIN csd.prop_stts prop_stts ON prop.prop_stts_code = prop_stts.prop_stts_code" & vbNewLine _
& "JOIN csd.natr_rqst natr_rqst ON natr_rqst.natr_rqst_code = prop.natr_rqst_code" & vbNewLine _
& "WHERE pi.prim_addr_flag='Y'" & vbNewLine _
& "ORDER BY lead, ILN DROP TABLE #myPids CREATE INDEX p1 ON #myProps(lead)" & vbNewLine
tempTables = tempTables & "SELECT p.lead, p.pi_last_name, rs.string, rs.score, rev_prop_vw.revr_id, rev_prop_vw.rev_rtrn_date, id=identity(18), 0 as 'seq'" & vbNewLine _
& "INTO #myRevs" & vbNewLine _
& "FROM #myProps p, csd.rev_prop_vw rev_prop_vw, tempdb.guest.revScores rs" & vbNewLine _
& "WHERE p.ILN < 'M' AND p.lead = rev_prop_vw.prop_id AND rev_prop_vw.rev_prop_rtng_ind = rs.yn " & vbNewLine _
& "ORDER BY lead, score DESC" & vbNewLine _
& "SELECT lead, MIN(id) as 'start' INTO #myStarts FROM #myRevs GROUP BY lead" & vbNewLine _
& "UPDATE #myRevs set seq = id-M.start FROM #myRevs r, #myStarts M WHERE r.lead = M.lead" & vbNewLine _
& "DROP TABLE #myStarts" & vbNewLine
mainQuery = "SELECT   rv.lead, rv.pi_last_name, 'Review' as docType, rev_prop.pm_logn_id, rv.revr_id as panl_revr_id,  revr.revr_last_name as 'name', revr_opt_addr_line.revr_addr_txt as 'info'," & vbNewLine _
& "convert(varchar,rev_prop.rev_type_code) AS type, convert(varchar,rev_prop.rev_stts_code) as stts, rev_prop.rev_due_date as due, rv.rev_rtrn_date as returned," & vbNewLine _
& " rv.string as score, rev_txt.REV_PROP_TXT_FLDS as 'text'" & vbNewLine _
& "FROM #myRevs rv, csd.revr revr, csd.rev_prop rev_prop, csd.revr_opt_addr_line revr_opt_addr_line, csd.rev_prop_txt_flds_vw rev_txt" & vbNewLine _
& "WHERE rv.lead = rev_prop.prop_id AND rv.lead = rev_txt.PROP_ID AND rv.revr_id = rev_prop.revr_id AND rv.revr_id = revr.revr_id AND rv.revr_id = revr_opt_addr_line.revr_id" & vbNewLine _
& "  AND rv.revr_id = rev_txt.REVR_ID AND ((revr_opt_addr_line.addr_lne_type_code='E'))" & vbNewLine _
& "UNION ALL SELECT p.lead, p.pi_last_name, ' PanlSumm', panl.pm_logn_id, panl_prop_summ.PANL_ID,  ' '+panl.panl_name, panl_rcom_def.RCOM_TXT," & vbNewLine _
& "convert(varchar,panl_prop_summ.RCOM_SEQ_NUM), convert(varchar,panl_prop_summ.PROP_ORDR), panl.panl_bgn_date, panl_prop_summ.panl_summ_rlse_date," & vbNewLine _
& "panl_rcom_def.RCOM_ABBR , panl_prop_summ.PANL_SUMM_TXT" & vbNewLine _
& "FROM #myProps p, FLflpdb.flp.panl_prop_summ panl_prop_summ, FLflpdb.flp.panl_rcom_def panl_rcom_def, csd.panl panl" & vbNewLine _
& "WHERE  p.ILN < 'M' AND p.lead = panl_prop_summ.PROP_ID AND panl_prop_summ.PANL_ID = panl.panl_id" & vbNewLine _
& "  AND panl_prop_summ.RCOM_SEQ_NUM = panl_rcom_def.RCOM_SEQ_NUM AND panl_prop_summ.PANL_ID = panl_rcom_def.PANL_ID" & vbNewLine _
& "UNION ALL SELECT p.lead, p.pi_last_name, 'POCmnt', p.PO, p.prop_id, cmt.cmnt_cre_id, '', '', convert(varchar,cmt.cmnt_prop_stts_code), cmt.beg_eff_date, cmt.end_eff_date,'', cmt.cmnt" & vbNewLine _
& "FROM #myProps p, FLflpdb.flp.cmnt_prop cmt WHERE p.prop_id = cmt.prop_id AND (p.ILN < 'M' OR LEN(cmt.cmnt) <> (SELECT LEN(l.cmnt) FROM FLflpdb.flp.cmnt_prop l WHERE p.lead = l.prop_id))" & vbNewLine _
& "UNION ALL SELECT p.lead, p.pi_last_name, 'RA' as docType, p.PO, '', ra.last_updt_user, '', '', null, null, ra.last_updt_tmsp, '', ra.prop_rev_anly_txt" & vbNewLine _
& "FROM #myProps p, csd.prop_rev_anly_vw ra WHERE p.ILN < 'M' AND p.lead = ra.prop_id" & vbNewLine
mainQuery = mainQuery & "UNION ALL SELECT p.lead, p.pi_last_name, 'Abstr', p.PO, '', a.last_updt_user, a.cent_mrkr_prop, a.cent_mrkr_awd, null, null, a.last_updt_tmsp,'', a.abst_narr_txt" & vbNewLine _
& "FROM #myProps p, csd.abst a WHERE p.prop_id = a.awd_id AND (p.ILN < 'M' OR LEN(a.abst_narr_txt) <>  (SELECT LEN(l.abst_narr_txt) FROM csd.abst l WHERE p.lead = l.awd_id))" & vbNewLine _
& "UNION ALL SELECT p.lead, p.pi_last_name, 'SummProj',p.PO, '', '', '', null, null, null, p.nsf_rcvd_date,'', s.summ" & vbNewLine _
& "FROM #myProps p, #mySumm s WHERE p.ILN < 'M' AND s.lead = p.lead" & vbNewLine _
& "" & vbNewLine _
& "ORDER BY lead, docType, revr_last_name, revr_id" & vbNewLine
dropTables = "DROP TABLE #myRevs,#myProps, #mySumm" & vbNewLine

    Set qt = Worksheets("ProjText").ListObjects.Item(1).QueryTable
    With qt
        .CommandText = setNC & myProps & tempTables & mainQuery & dropTables
        'MsgBox Mid(.CommandText, 1000, 1000)
        .Refresh
    End With
    
'Clear out any templates chosen because we are going to refresh the AllRACandidatesTable
Range("AllRACandidatesTable[RATemplate]").ClearContents
   
tempTables = "-- PRCs for props" & vbNewLine _
& "SELECT DISTINCT p.prop_id, pa.prop_atr_code, id=identity(18), 0 as seq" & vbNewLine _
& "INTO #myPRCs" & vbNewLine _
& "FROM #myPids p" & vbNewLine _
& "JOIN csd.prop_atr pa ON pa.prop_id = p.prop_id AND pa.prop_atr_type_code = 'PRC'" & vbNewLine _
& "ORDER BY p.prop_id, pa.prop_atr_code" & vbNewLine _
& "SELECT prop_id, MIN(id) as start INTO #myPRCStart FROM #myPRCs GROUP BY prop_id" & vbNewLine _
& "UPDATE #myPRCs set seq = id-M.start FROM #myPRCs r, #myPRCStart M WHERE r.prop_id = M.prop_id" & vbNewLine _
& "DROP TABLE #myPRCStart" & vbNewLine _
& "declare @PRCseparator varchar(2)" & vbNewLine _
& "set @PRCseparator = ', ' --char(10)+char(13)" & vbNewLine _
& "-- per-proposal data" & vbNewLine _
& "SELECT pid.lead, pid.ILN, pid.prop_id," & vbNewLine _
& "pi_last_name as L, pi_frst_name as F, inst_shrt_name as I, pi_emai_addr AS M, rqst_dol as D, " & vbNewLine _
& "prc.PRCs as R, prc2.PRCs AS RN, b.revn_num as N, b.Tot as T, id=identity(18), 0 as seq" & vbNewLine _
& "INTO #myProp" & vbNewLine _
& "FROM #myPids pid" & vbNewLine _
& "JOIN csd.prop p ON p.prop_id = pid.prop_id" & vbNewLine _
& "LEFT OUTER JOIN csd.pi_vw pi ON pi.pi_id = p.pi_id" & vbNewLine _
& "LEFT OUTER JOIN csd.inst inst ON inst.inst_id = p.inst_id" & vbNewLine _
& "LEFT OUTER JOIN (SELECT p.prop_id, eb.revn_num, SUM(eb.budg_tot_dol) as Tot--, COUNT(eb.budg_seq_yr) as Yrs" & vbNewLine _
& "    FROM #myPids p" & vbNewLine _
& "    JOIN csd.eps_blip eb ON p.prop_id = eb.prop_id AND NOT EXISTS (SELECT eb1.revn_num FROM csd.eps_blip eb1 WHERE eb.prop_id = eb1.prop_id AND eb.revn_num < eb1.revn_num)" & vbNewLine _
& "    GROUP BY p.prop_id, eb.revn_num) b ON b.prop_id = pid.prop_id" & vbNewLine _
& "LEFT OUTER JOIN (SELECT prop_id," & vbNewLine
tempTables = tempTables & "        MAX( CASE pa.seq WHEN 0 THEN       pa.prop_atr_code ELSE '' END ) +" & vbNewLine _
& "        MAX( CASE pa.seq WHEN 1 THEN ' ' + pa.prop_atr_code ELSE '' END ) +" & vbNewLine _
& "        MAX( CASE pa.seq WHEN 2 THEN ' ' + pa.prop_atr_code ELSE '' END ) +" & vbNewLine _
& "        MAX( CASE pa.seq WHEN 3 THEN ' ' + pa.prop_atr_code ELSE '' END ) +" & vbNewLine _
& "        MAX( CASE pa.seq WHEN 4 THEN ' ' + pa.prop_atr_code ELSE '' END ) +" & vbNewLine _
& "        MAX( CASE pa.seq WHEN 5 THEN ' ' + pa.prop_atr_code ELSE '' END ) +" & vbNewLine _
& "        MAX( CASE pa.seq WHEN 6 THEN ' ' + pa.prop_atr_code ELSE '' END ) AS PRCs" & vbNewLine _
& "    FROM #myPRCs pa " & vbNewLine _
& "    GROUP BY prop_id) prc ON prc.prop_id = pid.prop_id" & vbNewLine _
& "LEFT OUTER JOIN (SELECT prop_id, MAX( CASE pa.seq WHEN 0 THEN rtrim(pa.prop_atr_code) + ': '+ pgm_ref_long_name ELSE '' END ) +" & vbNewLine _
& "        MAX( CASE pa.seq WHEN 1 THEN @PRCseparator + rtrim(pa.prop_atr_code) + ': '+ pgm_ref_long_name ELSE '' END ) +" & vbNewLine _
& "        MAX( CASE pa.seq WHEN 2 THEN @PRCseparator + rtrim(pa.prop_atr_code) + ': '+ pgm_ref_long_name ELSE '' END ) +" & vbNewLine _
& "        MAX( CASE pa.seq WHEN 3 THEN @PRCseparator + rtrim(pa.prop_atr_code) + ': '+ pgm_ref_long_name ELSE '' END ) +" & vbNewLine _
& "        MAX( CASE pa.seq WHEN 4 THEN @PRCseparator + rtrim(pa.prop_atr_code) + ': '+ pgm_ref_long_name ELSE '' END ) +" & vbNewLine _
& "        MAX( CASE pa.seq WHEN 5 THEN @PRCseparator + rtrim(pa.prop_atr_code) + ': '+ pgm_ref_long_name ELSE '' END ) +" & vbNewLine _
& "        MAX( CASE pa.seq WHEN 6 THEN @PRCseparator + rtrim(pa.prop_atr_code) + ': '+ pgm_ref_long_name ELSE '' END ) AS 'PRCs'" & vbNewLine _
& "    FROM #myPRCs pa " & vbNewLine _
& "    JOIN csd.pgm_ref prc ON pa.prop_atr_code = prc.pgm_ref_code" & vbNewLine _
& "    GROUP BY prop_id) prc2 ON prc2.prop_id = pid.prop_id" & vbNewLine _
& "WHERE pi.prim_addr_flag='Y'" & vbNewLine _
& "ORDER BY lead, ILN, prop_id" & vbNewLine _
& "DROP TABLE #myPRCs" & vbNewLine
tempTables = tempTables & "SELECT lead, MIN(id) as start INTO #myPropStart FROM #myProp GROUP BY lead" & vbNewLine _
& "UPDATE #myProp set seq = id-M.start FROM #myProp r, #myPropStart M WHERE r.lead = M.lead" & vbNewLine _
& "DROP TABLE #myPropStart" & vbNewLine _
& "SELECT rp.*, " & vbNewLine _
& "CASE WHEN rpv.rev_prop_unrl_flag  = 'Y' THEN 1 ELSE 0 END as unrlsbl," & vbNewLine _
& "CASE WHEN (rpv.rev_rlse_flag  <> 'Y'AND rpv.rev_prop_unrl_flag  <> 'Y') THEN 1 ELSE 0 END as unmkd, " & vbNewLine _
& "rs.string, rs.score, id=identity(18), 0 as 'seq'" & vbNewLine _
& "INTO #myRevs FROM (" & vbNewLine _
& "    SELECT p.lead, pp.panl_id, rp.revr_id, rp.rev_stts_code, rp.rev_type_code, rp.rev_sent_date,rp.rev_rtrn_date, rp.rev_due_date" & vbNewLine _
& "    FROM #myPids p " & vbNewLine _
& "    JOIN csd.panl_prop pp ON pp.prop_id = p.lead" & vbNewLine _
& "    JOIN csd.rev_prop rp ON rp.prop_id = p.lead AND rp.rev_type_code<>'R'" & vbNewLine _
& "    JOIN csd.panl_revr pr ON pr.panl_id = pp.panl_id AND pr.revr_id = rp.revr_id" & vbNewLine _
& "    WHERE ILN < 'M'AND rp.rev_stts_code <> 'C'AND rev_rtrn_date is not null " & vbNewLine _
& "UNION ALL SELECT p.lead, '.ad hoc', rp.revr_id, rp.rev_stts_code, rp.rev_type_code, rp.rev_sent_date, rp.rev_rtrn_date, rp.rev_due_date" & vbNewLine _
& "    FROM #myPids p " & vbNewLine _
& "    JOIN csd.rev_prop rp ON rp.prop_id = p.lead AND rp.rev_type_code='R' -- ad hoc only" & vbNewLine _
& "    WHERE ILN < 'M'AND rp.rev_stts_code <> 'C' AND rev_rtrn_date is not null ) rp" & vbNewLine _
& "LEFT OUTER JOIN csd.rev_prop_vw rpv ON rpv.prop_id = rp.lead AND rpv.revr_id = rp.revr_id" & vbNewLine _
& "LEFT OUTER JOIN tempdb.guest.revScores rs ON rs.yn = rpv.rev_prop_rtng_ind" & vbNewLine _
& "ORDER BY lead, score DESC, panl_id " & vbNewLine _
& "SELECT lead, MIN(id) as 'start' INTO #myStarts FROM #myRevs GROUP BY lead" & vbNewLine _
& "UPDATE #myRevs set seq = id-M.start FROM #myRevs r, #myStarts M  WHERE r.lead = M.lead" & vbNewLine _
& "DROP TABLE #myStarts" & vbNewLine

tempTables = tempTables & "--r.Nrev,r.Nunrlsd,r.Nunmkd,r.reviews,r.avg_score,r.last_rev_date" & vbNewLine _
& "SELECT lead, count(revr_id) as Nrev, nullif(sum(unrlsbl),0) as Nunrlsd, nullif(sum(unmkd),0) as Nunmkd," & vbNewLine _
& " MAX(CASE r.seq WHEN  0 THEN r.string ELSE '' END)+" & vbNewLine _
& " MAX(CASE r.seq WHEN  1 THEN ','+r.string ELSE '' END)+" & vbNewLine _
& " MAX(CASE r.seq WHEN  2 THEN ','+r.string ELSE '' END)+" & vbNewLine _
& " MAX(CASE r.seq WHEN  3 THEN ','+r.string ELSE '' END)+" & vbNewLine _
& " MAX(CASE r.seq WHEN  4 THEN ','+r.string ELSE '' END)+" & vbNewLine _
& " MAX(CASE r.seq WHEN  5 THEN ','+r.string ELSE '' END)+" & vbNewLine _
& " MAX(CASE r.seq WHEN  6 THEN ','+r.string ELSE '' END)+" & vbNewLine _
& " MAX(CASE r.seq WHEN  7 THEN ','+r.string ELSE '' END)+" & vbNewLine _
& " MAX(CASE r.seq WHEN  8 THEN ','+r.string ELSE '' END)+" & vbNewLine _
& " MAX(CASE r.seq WHEN  9 THEN ','+r.string ELSE '' END)+" & vbNewLine _
& " MAX(CASE r.seq WHEN 10 THEN ','+r.string ELSE '' END)+" & vbNewLine _
& " MAX(CASE r.seq WHEN 11 THEN ','+r.string ELSE '' END)+" & vbNewLine _
& " MAX(CASE r.seq WHEN 12 THEN ','+r.string ELSE '' END) AS reviews, AVG(r.score) AS avg_score, MAX(r.rev_rtrn_date) AS last_rev_date" & vbNewLine _
& "INTO #myRevSumm" & vbNewLine _
& "FROM #myRevs r" & vbNewLine _
& "GROUP BY r.lead" & vbNewLine _
& "ORDER BY r.lead" & vbNewLine _
& "CREATE INDEX myRevSumm_ix ON #myRevSumm(lead)" & vbNewLine

tempTables = tempTables & "SELECT lead,panl_id, count(revr_id) as Nrev,STUFF(LTRIM(" & vbNewLine _
& " MAX(CASE r.seq WHEN  0 THEN ','+r.string ELSE '' END)+" & vbNewLine _
& " MAX(CASE r.seq WHEN  1 THEN ','+r.string ELSE '' END)+" & vbNewLine _
& " MAX(CASE r.seq WHEN  2 THEN ','+r.string ELSE '' END)+" & vbNewLine _
& " MAX(CASE r.seq WHEN  3 THEN ','+r.string ELSE '' END)+" & vbNewLine _
& " MAX(CASE r.seq WHEN  4 THEN ','+r.string ELSE '' END)+" & vbNewLine _
& " MAX(CASE r.seq WHEN  5 THEN ','+r.string ELSE '' END)+" & vbNewLine _
& " MAX(CASE r.seq WHEN  6 THEN ','+r.string ELSE '' END)+" & vbNewLine _
& " MAX(CASE r.seq WHEN  7 THEN ','+r.string ELSE '' END)+" & vbNewLine _
& " MAX(CASE r.seq WHEN  8 THEN ','+r.string ELSE '' END)+" & vbNewLine _
& " MAX(CASE r.seq WHEN  9 THEN ','+r.string ELSE '' END)+" & vbNewLine _
& " MAX(CASE r.seq WHEN 10 THEN ','+r.string ELSE '' END)+" & vbNewLine _
& " MAX(CASE r.seq WHEN 11 THEN ','+r.string ELSE '' END)+" & vbNewLine _
& " MAX(CASE r.seq WHEN 12 THEN ','+r.string ELSE '' END)),1,1,'') AS reviews, MAX(r.rev_rtrn_date) AS last_rev_date" & vbNewLine _
& "INTO #myRevPanl" & vbNewLine _
& "FROM #myRevs r" & vbNewLine _
& "GROUP BY lead, panl_id" & vbNewLine _
& "ORDER BY lead, panl_id" & vbNewLine _
& "-- panel summaries" & vbNewLine _
& "SELECT ps.panl_id as I,panl_name as N, panl_end_date as E, convert(varchar,SUM(ps.rtCount)) + ' rated projects: ' +" & vbNewLine _
& "        MAX( CASE ps.RCOM_SEQ_NUM WHEN 1 THEN        convert(varchar,ps.rtCount) + ' ' +  ps.RCOM_ABBR ELSE '' END ) +" & vbNewLine _
& "        MAX( CASE ps.RCOM_SEQ_NUM WHEN 2 THEN ', ' + convert(varchar,ps.rtCount) + ' ' +  ps.RCOM_ABBR ELSE '' END ) +" & vbNewLine _
& "        MAX( CASE ps.RCOM_SEQ_NUM WHEN 3 THEN ', ' + convert(varchar,ps.rtCount) + ' ' +  ps.RCOM_ABBR ELSE '' END ) +" & vbNewLine _
& "        MAX( CASE ps.RCOM_SEQ_NUM WHEN 4 THEN ', ' + convert(varchar,ps.rtCount) + ' ' +  ps.RCOM_ABBR ELSE '' END ) +" & vbNewLine
tempTables = tempTables & "        MAX( CASE ps.RCOM_SEQ_NUM WHEN 5 THEN ', ' + convert(varchar,ps.rtCount) + ' ' +  ps.RCOM_ABBR ELSE '' END ) +" & vbNewLine _
& "        MAX( CASE ps.RCOM_SEQ_NUM WHEN 6 THEN ', ' + convert(varchar,ps.rtCount) + ' ' +  ps.RCOM_ABBR ELSE '' END ) as S" & vbNewLine _
& "INTO #panlSumm" & vbNewLine _
& "FROM  csd.panl panl," & vbNewLine _
& "      (SELECT pl.panl_id, prd.RCOM_SEQ_NUM, prd.RCOM_ABBR, Count(pps.PROP_ID) AS rtCount" & vbNewLine _
& "        FROM FLflpdb.flp.panl_prop_summ pps, FLflpdb.flp.panl_rcom_def prd, csd.prop pr," & vbNewLine _
& "             (SELECT DISTINCT panl_id FROM #myPids p " & vbNewLine _
& "              JOIN csd.panl_prop pp ON pp.prop_id = p.prop_id) pl" & vbNewLine _
& "        WHERE pl.panl_id = pps.PANL_ID   ------- panel must have one of my proposals" & vbNewLine _
& "          AND pps.PROP_ID = pr.prop_id AND pr.prop_id=isnull(pr.lead_prop_id,pr.prop_id) --- count leads only" & vbNewLine _
& "          AND pps.PANL_ID *= prd.PANL_ID AND pps.RCOM_SEQ_NUM *= prd.RCOM_SEQ_NUM" & vbNewLine _
& "       GROUP BY pl.panl_id, prd.RCOM_SEQ_NUM, prd.RCOM_ABBR ) ps" & vbNewLine _
& "WHERE ps.PANL_ID = panl.PANL_ID" & vbNewLine _
& "GROUP BY ps.panl_id, panl.panl_name, panl_end_date" & vbNewLine _
& "ORDER BY ps.panl_id" & vbNewLine

tempTables = tempTables & "SELECT rp.lead,  PS.*, pps.RCOM_SEQ_NUM AS RS, prd.RCOM_ABBR as RA, prd.RCOM_TXT as RT, pps.PROP_ORDR as RK," & vbNewLine _
& "rp.reviews as V,rp.last_rev_date as D, -- pps.PANL_SUMM_TXT as T," & vbNewLine _
& "CASE WHEN panl_summ_rlse_flag = 'Y' THEN 1 ELSE 0 END as summ_rlse, " & vbNewLine _
& "CASE WHEN (panl_summ_unrl_flag <> 'Y' AND panl_summ_rlse_flag <> 'Y')  THEN 1 ELSE 0 END as summ_unmrkd," & vbNewLine _
& "id=identity(18), 0 as seq" & vbNewLine _
& "INTO #myProjPanl" & vbNewLine _
& "FROM #myRevPanl rp" & vbNewLine _
& "JOIN #PanlSumm PS ON PS.I = rp.panl_id" & vbNewLine _
& "LEFT OUTER JOIN FLflpdb.flp.panl_prop_summ pps ON pps.panl_id = rp.panl_id AND pps.prop_id = lead" & vbNewLine _
& "LEFT OUTER JOIN FLflpdb.flp.panl_rcom_def prd ON prd.panl_id = rp.panl_id  AND prd.RCOM_SEQ_NUM = pps.RCOM_SEQ_NUM" & vbNewLine _
& "order by lead, PS.E" & vbNewLine _
& "SELECT lead, MIN(id) as 'start' INTO #myPStarts FROM #myProjPanl GROUP BY lead" & vbNewLine _
& "UPDATE #myProjPanl set seq = id-M.start FROM #myProjPanl r, #myPStarts M  WHERE r.lead = M.lead" & vbNewLine _
& "DROP TABLE #myPStarts" & vbNewLine

mainQuery = "declare @olddate datetime" & vbNewLine _
& "set @olddate = '1/1/2000'" & vbNewLine _
& "SELECT getdate() AS pulldate, nsf_rcvd_date, ra.last_updt_tmsp as RAupdate, nullif(dd_rcom_date,'1900-01-01') AS dd_rcom_date, cntx_stmt_id," & vbNewLine _
& "prop.pgm_annc_id, prop.org_code, prop.pgm_ele_code, prop.pm_ibm_logn_id as PO,prop_stts_abbr,natr_rqst.natr_rqst_abbr,prop.obj_clas_code," & vbNewLine _
& "isnull(RecRkMin,9) AS RecRkMin, pn0.RA AS RCOM_ABBR0,pn1.RA AS RCOM_ABBR1, pn2.RA AS RCOM_ABBR2," & vbNewLine _
& "r.reviews,r.avg_score," & vbNewLine _
& "r.Nrev,r.Nunrlsd,r.Nunmkd,r.last_rev_date," & vbNewLine _
& "nPanl, nPSrlse, nPSunmrkd," & vbNewLine _
& "org.dir_div_abbr as Div, p0.prop_id as prop_id0, p0.L as last0, p0.F as frst0, p0.I as inst0, p0.D as rqst0, p0.T as b0tot," & vbNewLine _
& "a.Nrev as AhNrev, a.reviews as AhRevs, a.last_rev_date as AhLast, " & vbNewLine _
& "   pn0.I AS panl_id0, pn0.RT AS RCOM_TXT0, pn0.RK AS rank0, pn0.N AS panl_name0,pn0.E AS panl_end0,pn0.V AS revs0,pn0.S AS panlSumm0,--pn0.T AS PStxt0," & vbNewLine _
& "   pn1.I AS panl_id1, pn1.RT AS RCOM_TXT1, pn1.RK AS rank1, pn1.N AS panl_name1,pn1.E AS panl_end1,pn1.V AS revs1,pn1.S AS panlSumm1,--pn1.T AS PStxt1," & vbNewLine _
& "   pn2.I AS panl_id2, pn2.RT AS RCOM_TXT2, pn2.RK AS rank2, pn2.N AS panl_name2,pn2.E AS panl_end2,pn2.V AS revs2,pn2.S AS panlSumm2,--pn2.T AS PStxt2," & vbNewLine _
& "rtrim(prop_titl_txt) AS prop_titl_txt, projTot.rqst_tot,budg_tot,budRevnMax, prop.rqst_eff_date, prop.rqst_mnth_cnt, " & vbNewLine _
& "rtrim(pa.dflt_prop_titl_txt) AS solicitation, org.org_long_name, " & vbNewLine _
& "pgm_ele_name, sign_blck_name, prop_stts.prop_stts_txt, obj_clas_name," & vbNewLine _
& "o2.dir_div_abbr as Dir, o2.org_long_name as Dir_name, " & vbNewLine _
& "p0.R as PRC0,p0.RN as PRCN0," & vbNewLine _
& "p1.prop_id as prop_id1, p1.L as last1, p1.F as frst1, p1.I as inst1, p1.D as rqst1, p1.T as b1tot, p1.R as PRC1,p1.RN as PRCN1," & vbNewLine _
& "p2.prop_id as prop_id2, p2.L as last2, p2.F as frst2, p2.I as inst2, p2.D as rqst2, p2.T as b2tot, p2.R as PRC2,p2.RN as PRCN2," & vbNewLine _
& "p3.prop_id as prop_id3, p3.L as last3, p3.F as frst3, p3.I as inst3, p3.D as rqst3, p3.T as b3tot, p3.R as PRC3,p3.RN as PRCN3," & vbNewLine _
& "p4.prop_id as prop_id4, p4.L as last4, p4.F as frst4, p4.I as inst4, p4.D as rqst4, p4.T as b4tot, p4.R as PRC4,p4.RN as PRCN4," & vbNewLine _
& "p5.prop_id as prop_id5, p5.L as last5, p5.F as frst5, p5.I as inst5, p5.D as rqst5, p5.T as b5tot, p5.R as PRC5,p5.RN as PRCN5," & vbNewLine _
& "p6.prop_id as prop_id6, p6.L as last6, p6.F as frst6, p6.I as inst6, p6.D as rqst6, p6.T as b6tot, p6.R as PRC6,p6.RN as PRCN6," & vbNewLine
mainQuery = mainQuery & "p0.M AS email, (SELECT MAX(CASE r.seq WHEN  0 THEN r.M ELSE '' END)+" & vbNewLine _
& " MAX(CASE r.seq WHEN  1 THEN ';'+r.M ELSE '' END)+" & vbNewLine _
& " MAX(CASE r.seq WHEN  2 THEN ';'+r.M ELSE '' END)+" & vbNewLine _
& " MAX(CASE r.seq WHEN  3 THEN ';'+r.M ELSE '' END)+" & vbNewLine _
& " MAX(CASE r.seq WHEN  4 THEN ';'+r.M ELSE '' END)+" & vbNewLine _
& " MAX(CASE r.seq WHEN  5 THEN ';'+r.M ELSE '' END)+" & vbNewLine _
& " MAX(CASE r.seq WHEN  6 THEN ';'+r.M ELSE '' END) " & vbNewLine _
& "FROM #myProp r WHERE r.lead = p.lead) AS allPIemail" & vbNewLine _
& "FROM #myPids p" & vbNewLine _
& "JOIN csd.prop prop ON prop.prop_id = p.lead" & vbNewLine _
& "JOIN csd.org org ON org.org_code = prop.org_code" & vbNewLine _
& "JOIN csd.org o2 ON o2.org_code =left(prop.org_code,2)+'000000' " & vbNewLine _
& "LEFT OUTER JOIN csd.pgm_ele pe ON pe.pgm_ele_code = prop.pgm_ele_code" & vbNewLine _
& "LEFT OUTER JOIN csd.pgm_annc pa ON pa.pgm_annc_id = prop.pgm_annc_id" & vbNewLine _
& "JOIN csd.obj_clas oc ON oc.obj_clas_code = prop.obj_clas_code" & vbNewLine _
& "JOIN csd.prop_stts prop_stts ON prop_stts.prop_stts_code = prop.prop_stts_code" & vbNewLine _
& "JOIN csd.natr_rqst natr_rqst ON natr_rqst.natr_rqst_code = prop.natr_rqst_code" & vbNewLine _
& "LEFT OUTER JOIN csd.prop_rev_anly_vw ra ON ra.prop_id = p.lead" & vbNewLine _
& "JOIN (SELECT lead, SUM(D) AS rqst_tot,MAX(N) AS budRevnMax, SUM(T) AS budg_tot " & vbNewLine _
& "    FROM #myProp GROUP BY lead) projTot ON projTot.lead = p.lead" & vbNewLine _
& "LEFT OUTER JOIN (SELECT * FROM #myProp mp WHERE mp.seq = 0) p0 ON p0.lead = p.lead" & vbNewLine _
& "LEFT OUTER JOIN (SELECT * FROM #myProp mp WHERE mp.seq = 1) p1 ON p1.lead = p.lead" & vbNewLine _
& "LEFT OUTER JOIN (SELECT * FROM #myProp mp WHERE mp.seq = 2) p2 ON p2.lead = p.lead" & vbNewLine _
& "LEFT OUTER JOIN (SELECT * FROM #myProp mp WHERE mp.seq = 3) p3 ON p3.lead = p.lead" & vbNewLine
mainQuery = mainQuery & "LEFT OUTER JOIN (SELECT * FROM #myProp mp WHERE mp.seq = 4) p4 ON p4.lead = p.lead" & vbNewLine _
& "LEFT OUTER JOIN (SELECT * FROM #myProp mp WHERE mp.seq = 5) p5 ON p5.lead = p.lead" & vbNewLine _
& "LEFT OUTER JOIN (SELECT * FROM #myProp mp WHERE mp.seq = 6) p6 ON p6.lead = p.lead" & vbNewLine _
& "LEFT OUTER JOIN #myRevSumm r ON r.lead = p.lead" & vbNewLine _
& "LEFT OUTER JOIN #myRevPanl a ON a.lead = p.lead AND panl_id = '.ad hoc'" & vbNewLine _
& "LEFT OUTER JOIN (SELECT lead, count(I) AS nPanl, min(RS)+isnull(min(RK),0)/100.0 AS RecRkMin, nullif(sum(summ_rlse),0) as nPSrlse, " & vbNewLine _
& "    nullif(sum(summ_unmrkd),0) as nPSunmrkd FROM #myProjPanl GROUP BY lead) pn ON pn.lead = p.lead" & vbNewLine _
& "LEFT OUTER JOIN (SELECT * FROM #myProjPanl WHERE 0=seq) pn0 ON pn0.lead = p.lead" & vbNewLine _
& "LEFT OUTER JOIN (SELECT * FROM #myProjPanl WHERE 1=seq) pn1 ON pn1.lead = p.lead" & vbNewLine _
& "LEFT OUTER JOIN (SELECT * FROM #myProjPanl WHERE 2=seq) pn2 ON pn2.lead = p.lead " & vbNewLine _
& "LEFT OUTER JOIN csd.po_vw po_vw ON po_vw.po_ibm_logn_id = prop.pm_ibm_logn_id" & vbNewLine _
& "WHERE p.ILN < 'M'" & vbNewLine _
& "UNION all SELECT getdate(),@olddate,@olddate,@olddate,'.first line.', -- dummy data to get MM formatting right" & vbNewLine _
& "'.dummy.','00000000','0000','.blank.','....','...','....', -1, -- min RecRkMin" & vbNewLine _
& "'.X','.X','.X', 'E,V,G,F,P',0, -- recommendations & revs" & vbNewLine _
& "0,0,0,@olddate, --revs" & vbNewLine _
& "0,0,0, -- nPanl" & vbNewLine _
& "'...','0000000', '.dummy.','.data.','.for formatting.', 0,0," & vbNewLine _
& "0,'E,V,G,F,P',@olddate, --ad hoc" & vbNewLine _
& "'P000000','.blank.',0,'.intentional blank.', @olddate, 'E,V,G,F,P', '.keep this line first.', --'.blank summary for formatting. Keep as first line. Mail-merge looks at the first few lines of the mail merge source to infer the types of data that it will see.  If those first few lines are null in a column, then it will guess that it is an number, which is the wrong thing to do for long text fields like this one. Therefore, I include a dummy first line that will never be used to generate an RA, but does have data of the correct format for MailMerge to pick up. '," & vbNewLine _
& "'P000001','.blank.',0,'.intentional blank.', @olddate, 'E,V,G,F,P', '.keep this line first.', --'.blank summary for formatting. Keep as first line. Mail-merge looks at the first few lines of the mail merge source to infer the types of data that it will see.  If those first few lines are null in a column, then it will guess that it is an number, which is the wrong thing to do for long text fields like this one. Therefore, I include a dummy first line that will never be used to generate an RA, but does have data of the correct format for MailMerge to pick up. '," & vbNewLine _
& "'P000002','.blank.',0,'.intentional blank.', @olddate, 'E,V,G,F,P', '.keep this line first.', --'.blank summary for formatting. Keep as first line. Mail-merge looks at the first few lines of the mail merge source to infer the types of data that it will see.  If those first few lines are null in a column, then it will guess that it is an number, which is the wrong thing to do for long text fields like this one. Therefore, I include a dummy first line that will never be used to generate an RA, but does have data of the correct format for MailMerge to pick up. '," & vbNewLine _
& "'.dummy title as first line for formatting.', 0,0,0, @olddate,0,'.dummy pgm annc/solic for formatting.', -- title-solic" & vbNewLine _
& "'.dummy div name.', '.dummy pec name.','Jack Snoeyink','.dummy status.','.dummy object class','NSF','.dummy dir name.'," & vbNewLine
mainQuery = mainQuery & " 'PRCs', 'PRCs Names'," & vbNewLine _
& "'0000000', '.dummy.','.data.','.for formatting.', 0,0, 'PRCs', 'PRCs Names'," & vbNewLine _
& "'0000000', '.dummy.','.data.','.for formatting.', 0,0, 'PRCs', 'PRCs Names'," & vbNewLine _
& "'0000000', '.dummy.','.data.','.for formatting.', 0,0, 'PRCs', 'PRCs Names'," & vbNewLine _
& "'0000000', '.dummy.','.data.','.for formatting.', 0,0, 'PRCs', 'PRCs Names'," & vbNewLine _
& "'0000000', '.dummy.','.data.','.for formatting.', 0,0, 'PRCs', 'PRCs Names'," & vbNewLine _
& "'0000000', '.dummy.','.data.','.for formatting.', 0,0, 'PRCs', 'PRCs Names'," & vbNewLine _
& "'.dummy email.','.dummy all-PIs email.'" & vbNewLine _
& "ORDER BY RecRkMin, p0.prop_id" & vbNewLine



dropTables = "DROP TABLE #myPids, #myProp, #myRevs, #myRevSumm, #myRevPanl, #myProjPanl, #PanlSumm" & vbNewLine


    
    Set qt = Worksheets("AllRACandidates").ListObjects.Item(1).QueryTable
    With qt
        .CommandText = setNC & myProps & tempTables & mainQuery & dropTables
        'MsgBox Mid(.CommandText, 1000, 1000)
        .Refresh (False)
    End With
Call CleanUpSheet(Worksheets("CheckCoding"))
Call CleanUpSheet(Worksheets("BudgetBlocks"))
Call CleanUpSheet(Worksheets("ProjText"))
Call CleanUpSheet(Worksheets("AllRACandidates"))
End Sub
