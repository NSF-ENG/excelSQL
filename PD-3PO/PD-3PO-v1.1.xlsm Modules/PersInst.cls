VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PersInst"
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
     .description = "Personnel & Insts from cover & budget pages"
     .tip = "Unhide to Jump to " & .name
     .tabtitle = "Personnel & Institutions"
     .note = "Personnel and institutions from cover pages and budget sheets." & vbNewLine & "Duplicates people listed with different names or institution spellings."
     .helpText = "Pulls PIs & co-PIs from cover pages, the Senior Personnel that recieve funds from budget pages, and any subaward lead.  It cannot retrieve unpaid personnel that co-manage subawards or collaborators. (See the SubAwards sheet to find the proposals with non-zero subaward amounts.) People will be duplicated if names or institutions are entered differently on, say, cover and budget pages."
     .helpvideo = ""
     .phase = 1
     .order = 200
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
Dim myPRCs, getPIs, getSrPers, mainQuery As String

'allPIs on these proposals
myPRCs = "SELECT DISTINCT prop.prop_id, pa.prop_atr_code, id=identity(18), 0 as 'seq' INTO #myPRCs" & vbNewLine _
& "FROM #myProps prop, csd.prop_atr pa WHERE pa.prop_id = prop.prop_id  AND pa.prop_atr_type_code = 'PRC'" & vbNewLine _
& "ORDER BY prop.prop_id, pa.prop_atr_code" & vbNewLine _
& "SELECT prop_id, MIN(id) as 'start' INTO #mySt2 FROM #myPRCs GROUP BY prop_id" & vbNewLine _
& "UPDATE #myPRCs set seq = id-M.start FROM #myPRCs r, #mySt2 M WHERE r.prop_id = M.prop_id" & vbNewLine _
& "DROP TABLE #mySt2" & vbNewLine

getPIs = "CREATE INDEX myPropsPIx ON #myProps(prop_id)" & vbNewLine & vbNewLine _
 & "SELECT prop.prop_id, prop.ILN, prop.prop_titl_txt, ' PI' as pers, pi.pi_last_name, pi.pi_frst_name, prop.inst_name,  pi.pi_gend_code, pi.pi_emai_addr, pi.pi_dept_name, prop.st_code, prop.inst_id, pi.pi_id, prop.rqst_dol" & vbNewLine _
 & "INTO #myPers " & vbNewLine _
 & "FROM #myProps prop, csd.pi_vw pi" & vbNewLine _
 & "WHERE prop.pi_id = pi.pi_id AND (pi.prim_addr_flag='Y') " & vbNewLine _
 & "UNION ALL SELECT prop.prop_id, prop.ILN, '', 'coPI', pi.pi_last_name, pi.pi_frst_name, inst.inst_shrt_name,  pi.pi_gend_code, pi.pi_emai_addr, pi.pi_dept_name, pi.st_code, pi.inst_id, pi.pi_id, 0" & vbNewLine _
 & "FROM csd.addl_pi_invl addl_pi_invl, csd.inst inst, csd.pi_vw pi, #myProps prop" & vbNewLine _
 & "  WHERE prop.prop_id = addl_pi_invl.prop_id AND addl_pi_invl.pi_id = pi.pi_id AND pi.inst_id = inst.inst_id AND (pi.prim_addr_flag='Y') " & vbNewLine _
 & "CREATE INDEX myPersPIx  ON #myPers(prop_id)" & vbNewLine & vbNewLine

getSrPers = "INSERT #myPers (prop_id, ILN, prop_titl_txt, pers, pi_last_name, pi_frst_name, inst_name, pi_gend_code, pi_emai_addr, pi_dept_name, st_code, inst_id, pi_id, rqst_dol)" & vbNewLine _
 & "SELECT DISTINCT prop.prop_id, prop.ILN, '', 'srPers', s.SR_LAST_NAME AS pi_last_name, s.SR_FRST_NAME AS pi_frst_name, inst.inst_shrt_name AS inst_name, '', '', '', inst.st_code, inst.inst_id, '', 0" & vbNewLine _
 & "FROM #myProps prop" & vbNewLine _
 & "JOIN csd.prop_subm_ctl_vw psc ON prop.prop_id = psc.PROP_ID" & vbNewLine _
 & "JOIN csd.sr_pers_resc_vw s ON psc.TEMP_PROP_ID = s.TEMP_PROP_ID" & vbNewLine _
 & "JOIN csd.inst inst ON s.PERF_INST_ID = inst.inst_id" & vbNewLine _
 & "  WHERE NOT EXISTS (SELECT * FROM #myPers p" & vbNewLine _
 & "WHERE prop.prop_id = p.prop_id  AND s.SR_LAST_NAME = p.pi_last_name AND s.SR_FRST_NAME = pi_frst_name AND inst.inst_shrt_name = p.inst_name)" & vbNewLine _
 & "INSERT #myPers (prop_id, ILN, prop_titl_txt, pers, pi_last_name, pi_frst_name, inst_name, pi_gend_code, pi_emai_addr, pi_dept_name, st_code, inst_id, pi_id, rqst_dol)" & vbNewLine _
 & "SELECT DISTINCT prop.prop_id, prop.ILN, '', 'subAwd', s.PI_LAST_NAME AS pi_last_name, s.PI_FRST_NAME AS pi_frst_name, inst.inst_shrt_name AS inst_name, '', '', '', inst.st_code, inst.inst_id, '', 0" & vbNewLine _
 & "FROM #myProps prop" & vbNewLine _
 & "JOIN csd.prop_subm_ctl_vw psc ON prop.prop_id = psc.PROP_ID" & vbNewLine _
 & "JOIN csd.subconpi_vw s ON psc.TEMP_PROP_ID = s.TEMP_PROP_ID" & vbNewLine _
 & "JOIN csd.inst inst ON s.PERF_INST_ID = inst.inst_id" & vbNewLine _
 & " WHERE NOT EXISTS (SELECT * FROM #myPers p" & vbNewLine _
 & "   WHERE prop.prop_id = p.prop_id  AND s.PI_LAST_NAME = p.pi_last_name AND s.PI_FRST_NAME = pi_frst_name AND inst.inst_shrt_name = p.inst_name)"

mainQuery = "SELECT " & limit & " prop.nsf_rcvd_date, prop.dd_rcom_date, prop.pgm_annc_id, prop.Dir, prop.org_code, prop.Pgm, prop.PO," & vbNewLine _
 & " prop.obj_clas_code, prop.natr_rqst_abbr, prop.prop_stts_abbr, prop.lead, prop.Div, pr.*, prop.rqst_mnth_cnt, prop.rqst_eff_date" & vbNewLine _
 & "FROM #myProps prop, #myPers pr" & vbNewLine _
 & "WHERE pr.prop_id = prop.prop_id" & vbNewLine _
 & "ORDER BY prop.lead, pr.ILN, pr.prop_id, pr.pers, pr.pi_last_name" & vbNewLine & vbNewLine

dropTables = "DROP TABLE #myPers  DROP TABLE #myProps"

Call QT.doQuery(mInputTab.InputSQL() & myPRCs & getPIs & getSrPers & mainQuery & dropTables)
End Sub




