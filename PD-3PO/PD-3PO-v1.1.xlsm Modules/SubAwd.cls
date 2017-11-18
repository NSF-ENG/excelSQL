VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SubAwd"
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
     .description = "Proposals with subawards in original budgets"
     .tip = "Unhide to jump to " & .name
     .tabtitle = "Subaward Details"
     .note = "Identifies proposals with non-zero subaward amounts in their budgets, and checks if subaward budgets match.  Rounding can cause small differences."
     .helpText = "Identifies proposals with non-zero subaward amounts in the original budgets, and checks that sub award amounts match. Has to recalculate indirect costs, so small discrepancies will be due to rounding. Large discrepancies are often missing budgets for subawards, or confusing consultants with subawards. Note: does not check budget revisions."
     .helpvideo = ""
     .phase = 3
     .order = 1401
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
Dim getSubAwds, makeBudg, Query As String

'Subawardee lists and dollar amount
getSubAwds = "SELECT p.prop_id, psc.temp_prop_id " & vbNewLine _
& "INTO #hasSubAwd" & vbNewLine _
& "FROM #myProps p " & vbNewLine _
& "JOIN csd.prop_subm_ctl_vw psc ON p.prop_id = psc.prop_id " & vbNewLine _
& "WHERE EXISTS (SELECT * FROM csd.budg_vw b WHERE psc.TEMP_PROP_ID = b.TEMP_PROP_ID AND b.sub_ctr_req_dol>0)  " & vbNewLine

makeBudg = "SELECT prop_id, perf_inst_id, revn_num, round(sum(subAwd),0) AS subAwd, round(sum(bud_tot),0) AS bud_tot  " & vbNewLine _
& "INTO #myBudg" & vbNewLine _
& "FROM ( " & vbNewLine _
& "  SELECT p.prop_id, b.perf_inst_id, b.revn_num, isnull(b.sub_ctr_req_dol,0) as subAwd, " & vbNewLine _
& "    isnull(OTHR_SR_REQ_DOL,0)+isnull(PDOC_REQ_DOL,0)+isnull(OTH_PROF_REQ_DOL,0)+isnull(GRAD_REQ_DOL,0)+isnull(UN_GRAD_REQ_DOL,0)+ " & vbNewLine _
& "    isnull(SEC_REQ_DOL,0)+isnull(OTH_PERS_REQ_DOL,0)+isnull(FRIN_BNFT_REQ_DOL,0)+isnull(DOM_TRAV_REQ_DOL,0)+isnull(FRGN_TRAV_REQ_DOL,0)+ " & vbNewLine _
& "    isnull(PART_SUPT_STPD_DOL,0)+isnull(PART_SUPT_TRAV_DOL,0)+isnull(PART_SUPT_SUBS_DOL,0)+isnull(PART_SUPT_OTH_DOL,0)+isnull(MATL_REQ_DOL,0)+ " & vbNewLine _
& "    isnull(PUB_REQ_DOL,0)+isnull(CNSL_REQ_DOL,0)+isnull(CPTR_SERV_REQ_DOL,0)+isnull(SUB_CTR_REQ_DOL,0)+isnull(OTH_DRCT_CST_REQ_DOL,0)+isnull(RSID_REQ_DOL,0) AS bud_tot " & vbNewLine _
& "   FROM #hasSubAwd p " & vbNewLine _
& "   JOIN csd.budg_vw b ON p.TEMP_PROP_ID = b.TEMP_PROP_ID " & vbNewLine

makeBudg = makeBudg & "  UNION ALL SELECT p.prop_id, sr.perf_inst_id, sr.revn_num, 0, sr_req_dol " & vbNewLine _
& "   FROM #hasSubAwd p" & vbNewLine _
& "   JOIN csd.sr_pers_resc_vw sr ON p.TEMP_PROP_ID = sr.TEMP_PROP_ID " & vbNewLine _
& "  UNION ALL SELECT p.prop_id, e.perf_inst_id, e.revn_num, 0, isnull(e.equp_cst_dol_req,0)" & vbNewLine _
& "   FROM #hasSubAwd p " & vbNewLine _
& "   JOIN csd.equp_cst_vw e ON p.TEMP_PROP_ID = e.TEMP_PROP_ID " & vbNewLine _
& "  UNION ALL SELECT p.prop_id, i.perf_inst_id, i.revn_num, 0, round(isnull(i.IDIR_CST_RATE*i.idir_cst_dol_req/100,0),0)" & vbNewLine _
& "   FROM #hasSubAwd p" & vbNewLine _
& "   JOIN csd.idir_cst_vw i ON p.TEMP_PROP_ID = i.TEMP_PROP_ID) b" & vbNewLine _
& "GROUP BY prop_id, perf_inst_id, revn_num  " & vbNewLine _
& "ORDER BY prop_id, subAwd DESC, bud_tot DESC " & vbNewLine _
& "CREATE INDEX myBudg_ix ON #myBudg(prop_id) " & vbNewLine

Query = "SELECT " & limit & " nsf_rcvd_date, dd_rcom_date, Dir, prop.org_code, Pgm, PO, prop.natr_rqst_abbr AS natr_rqst, prop.prop_stts_abbr AS prop_stts," & vbNewLine _
& "prop.ILN, prop.lead, Div, prop.prop_id, prop.pi_last_name, prop.inst_name," & vbNewLine _
& " x.inst_shrt_name AS perf_inst, x.revn_num, x.bud_tot AS budg_total, x.budg_diff, x.subAwd, sub_diff as subAwd_diff,  " & vbNewLine _
& "prop.prop_titl_txt" & vbNewLine _
& "FROM #myProps prop" & vbNewLine _
& "JOIN (SELECT prop_id, inst.inst_shrt_name, revn_num,  " & vbNewLine _
& "    bud_tot,  NULLIF(bud_tot - (SELECT SUM(b.budg_tot_dol) FROM csd.eps_blip b WHERE b.prop_id = p.prop_id AND b.revn_num = 0),0) AS budg_diff,  " & vbNewLine _
& "    subAwd, NULLIF(subAwd - (SELECT ISNULL(SUM(bud_tot),0) FROM #myBudg s WHERE subAwd = 0 AND s.prop_id = p.prop_id AND s.revn_num = p.revn_num),0) AS sub_diff  " & vbNewLine _
& "    FROM #myBudg p " & vbNewLine _
& "JOIN csd.inst inst ON p.perf_inst_id = inst.inst_id " & vbNewLine _
& "WHERE subAwd > 0) x  ON prop.prop_id = x.prop_id" & vbNewLine _
& "ORDER BY prop.lead, prop.ILN, prop.prop_id" & vbNewLine

dropTables = "DROP TABLE #myProps DROP TABLE #hasSubAwd DROP TABLE #myBudg" & vbNewLine
 
Call QT.doQuery(mInputTab.InputSQL & getSubAwds & makeBudg & Query & dropTables)
End Sub
