Attribute VB_Name = "Module1"

Dim selectProps As String

Sub propsFromCriteria()
Call Show
  
Dim PgmAnnc As String
PgmAnnc = Range("pgm_annc")
Dim OrgCode As String
OrgCode = Range("org_code")
Dim PEC As String
PEC = Range("PEC")
Dim FromDate As String
FromDate = Format(Range("from_date"), "yyyy-mm-dd hh:mm:ss")
Dim ToDate As String
ToDate = Format(Range("to_date"), "yyyy-mm-dd hh:mm:ss")

 selectProps = "SELECT DISTINCT isnull(prop.lead_prop_id,prop.prop_id) as lead" & vbNewLine & _
  "INTO #myLeads -- get distinct project leads fm input list" & vbNewLine & _
  "FROM flp.prop_pars prop" & vbNewLine & _
  "WHERE ((prop.pgm_annc_id Like '" & PgmAnnc & "') AND (prop.org_code Like '" & OrgCode & "') AND (prop.pgm_ele_code Like '" & PEC & _
    "') AND (prop.nsf_rcvd_date Between {ts '" & FromDate & "'} And {ts DATEADD(day,1,'" & ToDate & "')}))" & vbNewLine & _
  "SELECT CASE WHEN prop.lead_prop_id IS NULL THEN 'I' ELSE 'L' END as ILN," & vbNewLine & _
  "ml.lead AS lead, prop.prop_id, c.TEMP_PROP_ID, prop.nsf_rcvd_date, prop.rqst_dol, prop.prop_titl_txt, prop.pi_id" & vbNewLine & _
  "INTO #myProjPropPI -- start with project leads" & vbNewLine & _
  "FROM #myLeads ml, flp.prop_pars prop, flp.prop_subm_ctl c" & vbNewLine & _
  "WHERE ml.lead = prop.prop_id AND prop.prop_id = c.PROP_ID -- add other filters here?" & vbNewLine & _
  "DROP TABLE #myLeads" & vbNewLine & _
  "INSERT INTO #myProjPropPI -- add all collabs" & vbNewLine & _
  "SELECT 'N' AS ILN, prop.lead_prop_id AS lead, prop.prop_id, c.TEMP_PROP_ID, prop.nsf_rcvd_date, prop.rqst_dol, prop.prop_titl_txt, prop.pi_id" & vbNewLine & _
  "FROM #myProjPropPI ppp, flp.prop_pars prop, flp.prop_subm_ctl c" & vbNewLine & _
  "WHERE ppp.PROP_ID = prop.lead_prop_id And ppp.PROP_ID <> prop.PROP_ID And prop.PROP_ID = c.PROP_ID" & vbNewLine
    
'MsgBox selectProps

Call Compliance_Query
End Sub

Sub propsFromList()
Call Show
 
Dim ids As String
Dim nrows As Long
nrows = Range("Input[prop_id]").Rows.COUNT
If nrows < 2 Then
   If nrows < 1 Then
     MsgBox "No proposals listed under prop_id column; aborting"
     Exit Sub
   Else
     ids = Range("Input[prop_id]").Value
   End If
Else ' we have at least two
  ids = Join(Application.Transpose(Range("Input[prop_id]").Value), "','")
End If
' strip spaces (visible and invisible) and ,'' from string.
ids = "(prop.prop_id In " & Replace(Replace(Replace("('" & ids & "'))", " ", ""), Chr(63), ""), ",''", "") & vbNewLine

'MsgBox "please check your ides : " + ids

selectProps = "SELECT DISTINCT isnull(prop.lead_prop_id,prop.prop_id) as lead" & vbNewLine & _
  "INTO #myLeads -- get distinct project leads fm input list" & vbNewLine & _
  "FROM flp.prop_pars prop" & vbNewLine & _
        "WHERE " & ids & vbNewLine & _
  "SELECT CASE WHEN prop.lead_prop_id IS NULL THEN 'I' ELSE 'L' END as ILN," & vbNewLine & _
  "ml.lead AS lead, prop.prop_id, c.TEMP_PROP_ID, prop.nsf_rcvd_date, prop.rqst_dol, prop.prop_titl_txt, prop.pi_id" & vbNewLine & _
  "INTO #myProjPropPI -- start with project leads" & vbNewLine & _
  "FROM #myLeads ml, flp.prop_pars prop, flp.prop_subm_ctl c" & vbNewLine & _
  "WHERE ml.lead = prop.prop_id AND prop.prop_id = c.PROP_ID -- add other filters here?" & vbNewLine & _
  "DROP TABLE #myLeads" & vbNewLine & _
  "INSERT INTO #myProjPropPI -- add all collabs" & vbNewLine & _
  "SELECT 'N' AS ILN, prop.lead_prop_id AS lead, prop.prop_id, c.TEMP_PROP_ID, prop.nsf_rcvd_date, prop.rqst_dol, prop.prop_titl_txt, prop.pi_id" & vbNewLine & _
  "FROM #myProjPropPI ppp, flp.prop_pars prop, flp.prop_subm_ctl c" & vbNewLine & _
  "WHERE ppp.PROP_ID = prop.lead_prop_id And ppp.PROP_ID <> prop.PROP_ID And prop.PROP_ID = c.PROP_ID" & vbNewLine
'MsgBox selectProps

Call Compliance_Query
End Sub

Sub propsFromAplusB()
Call Show


'-- get prop by program info
Dim PgmAnnc As String
PgmAnnc = Range("pgm_annc")
Dim OrgCode As String
OrgCode = Range("org_code")
Dim PEC As String
PEC = Range("PEC")
Dim FromDate As String
FromDate = Format(Range("from_date"), "yyyy-mm-dd hh:mm:ss")
Dim ToDate As String
ToDate = Format(Range("to_date"), "yyyy-mm-dd hh:mm:ss")

'-- get prop by list
Dim ids As String
Dim nrows As Long
nrows = Range("Input[prop_id]").Rows.COUNT
If nrows < 2 Then
   If nrows < 1 Then
     MsgBox "No proposals listed under prop_id column; aborting"
     Exit Sub
   Else
     ids = Range("Input[prop_id]").Value
   End If
Else ' we have at least two
  ids = Join(Application.Transpose(Range("Input[prop_id]").Value), "','")
End If
' strip spaces (visible and invisible) and ,'' from string.
ids = "(prop.prop_id In " & Replace(Replace(Replace("('" & ids & "'))", " ", ""), Chr(63), ""), ",''", "") & vbNewLine
'MsgBox "please check your ides : " + ids


 selectProps = "SELECT DISTINCT isnull(prop.lead_prop_id,prop.prop_id) as lead" & vbNewLine & _
  "INTO #myLeads -- get distinct project leads fm input list" & vbNewLine & _
  "FROM flp.prop_pars prop" & vbNewLine & _
  "WHERE ((prop.pgm_annc_id Like '" & PgmAnnc & "') AND (prop.org_code Like '" & OrgCode & "') AND (prop.pgm_ele_code Like '" & PEC & _
    "') AND (prop.nsf_rcvd_date Between {ts '" & FromDate & "'} And {ts DATEADD(day,1,'" & ToDate & "')}))" & vbNewLine & _
    "OR " & ids & vbNewLine & _
  "SELECT CASE WHEN prop.lead_prop_id IS NULL THEN 'I' ELSE 'L' END as ILN," & vbNewLine & _
  "ml.lead AS lead, prop.prop_id, c.TEMP_PROP_ID, prop.nsf_rcvd_date, prop.rqst_dol, prop.prop_titl_txt, prop.pi_id" & vbNewLine & _
  "INTO #myProjPropPI -- start with project leads" & vbNewLine & _
  "FROM #myLeads ml, flp.prop_pars prop, flp.prop_subm_ctl c" & vbNewLine & _
  "WHERE ml.lead = prop.prop_id AND prop.prop_id = c.PROP_ID -- add other filters here?" & vbNewLine & _
  "DROP TABLE #myLeads" & vbNewLine & _
  "INSERT INTO #myProjPropPI -- add all collabs" & vbNewLine & _
  "SELECT 'N' AS ILN, prop.lead_prop_id AS lead, prop.prop_id, c.TEMP_PROP_ID, prop.nsf_rcvd_date, prop.rqst_dol, prop.prop_titl_txt, prop.pi_id" & vbNewLine & _
  "FROM #myProjPropPI ppp, flp.prop_pars prop, flp.prop_subm_ctl c" & vbNewLine & _
  "WHERE ppp.PROP_ID = prop.lead_prop_id And ppp.PROP_ID <> prop.PROP_ID And prop.PROP_ID = c.PROP_ID" & vbNewLine
    
'MsgBox selectProps

Call Compliance_Query
End Sub

Sub propsFromAminusB()
Call Show
'-- get prop by program info
Dim PgmAnnc As String
PgmAnnc = Range("pgm_annc")
Dim OrgCode As String
OrgCode = Range("org_code")
Dim PEC As String
PEC = Range("PEC")
Dim FromDate As String
FromDate = Format(Range("from_date"), "yyyy-mm-dd hh:mm:ss")
Dim ToDate As String
ToDate = Format(Range("to_date"), "yyyy-mm-dd hh:mm:ss")

'-- get prop by list
Dim ids As String
Dim nrows As Long
nrows = Range("Input[prop_id]").Rows.COUNT
If nrows < 2 Then
   If nrows < 1 Then
     MsgBox "No proposals listed under prop_id column; aborting"
     Exit Sub
   Else
     ids = Range("Input[prop_id]").Value
   End If
Else ' we have at least two
  ids = Join(Application.Transpose(Range("Input[prop_id]").Value), "','")
End If
' strip spaces (visible and invisible) and ,'' from string.
ids = "(prop.prop_id NOT In " & Replace(Replace(Replace("('" & ids & "'))", " ", ""), Chr(63), ""), ",''", "") & vbNewLine
'MsgBox "please check your ids : " + ids


 selectProps = "SELECT DISTINCT isnull(prop.lead_prop_id,prop.prop_id) as lead" & vbNewLine & _
  "INTO #myLeads -- get distinct project leads fm input list" & vbNewLine & _
  "FROM flp.prop_pars prop" & vbNewLine & _
  "WHERE ((prop.pgm_annc_id Like '" & PgmAnnc & "') AND (prop.org_code Like '" & OrgCode & "') AND (prop.pgm_ele_code Like '" & PEC & _
    "') AND (prop.nsf_rcvd_date Between {ts '" & FromDate & "'} And {ts DATEADD(day,1,'" & ToDate & "')}))" & vbNewLine & _
    "AND " & ids & vbNewLine & _
  "SELECT CASE WHEN prop.lead_prop_id IS NULL THEN 'I' ELSE 'L' END as ILN," & vbNewLine & _
  "ml.lead AS lead, prop.prop_id, c.TEMP_PROP_ID, prop.nsf_rcvd_date, prop.rqst_dol, prop.prop_titl_txt, prop.pi_id" & vbNewLine & _
  "INTO #myProjPropPI -- start with project leads" & vbNewLine & _
  "FROM #myLeads ml, flp.prop_pars prop, flp.prop_subm_ctl c" & vbNewLine & _
  "WHERE ml.lead = prop.prop_id AND prop.prop_id = c.PROP_ID -- add other filters here?" & vbNewLine & _
  "DROP TABLE #myLeads" & vbNewLine & _
  "INSERT INTO #myProjPropPI -- add all collabs" & vbNewLine & _
  "SELECT 'N' AS ILN, prop.lead_prop_id AS lead, prop.prop_id, c.TEMP_PROP_ID, prop.nsf_rcvd_date, prop.rqst_dol, prop.prop_titl_txt, prop.pi_id" & vbNewLine & _
  "FROM #myProjPropPI ppp, flp.prop_pars prop, flp.prop_subm_ctl c" & vbNewLine & _
  "WHERE ppp.PROP_ID = prop.lead_prop_id And ppp.PROP_ID <> prop.PROP_ID And prop.PROP_ID = c.PROP_ID" & vbNewLine
    
'MsgBox selectProps

Call Compliance_Query
End Sub

Sub checkAccess()

Call handlePwdForm ' get password and stuff into connection strings for all queries

    Dim QT As QueryTable
    Set QT = Worksheets("Help").ListObjects.Item(1).QueryTable
    With QT
        .CommandText = "select dbu.name AS db_name, so.name AS tbl_name FROM FLflpdb.dbo.sysobjects so" & vbNewLine _
& "JOIN FLflpdb.dbo.sysusers dbu ON so.uid = dbu.uid" & vbNewLine _
& "WHERE dbu.name = 'flp' AND so.name in ('addl_pi_invl_pars', 'awd_pars', 'awd_pi_copi_pars', 'bibl', 'budg','BUDG_EXPL', 'clbr_affl', 'DMP_PLAN_DOC', 'FAC_EQUP', 'ggpi_trck', 'inst', 'MENT_PLAN_DOC', 'pi', 'pi_copi_char', 'PI_COPI_PRIR_SUPT', 'PROJ_DESC', 'proj_summ', 'PROP_COVR', 'prop_pars', 'PROP_SPCL_ITEM', 'prop_stts_pars', 'prop_subm_ctl', 'supp_dtls', 'tz_ctry_map')" & vbNewLine _
& "and not exists (SELECT su.name FROM FLflpdb.dbo.sysprotects sp" & vbNewLine _
& " JOIN FLflpdb.dbo.sysusers su on sp.uid = su.uid" & vbNewLine _
& " WHERE sp.id = so.id and sp.action = 193 and su.name in ('public','" & userId & "'))" & vbNewLine _
& "UNION ALL SELECT dbu.name, so.name FROM sysobjects so " & vbNewLine _
& "JOIN sysusers dbu ON so.uid = dbu.uid" & vbNewLine _
& "WHERE dbu.name = 'csd' AND so.name in ('natr_rqst','eps_blip')" & vbNewLine _
& "and not exists (SELECT su.name FROM sysprotects sp" & vbNewLine _
& " JOIN sysusers su on sp.uid = su.uid" & vbNewLine _
& " WHERE sp.id = so.id and sp.action = 193 and su.name in ('public','" & userId & "'))" & vbNewLine _
& "order by dbu.name, so.name" & vbNewLine
        .Refresh BackgroundQuery:=False
    End With
    
End Sub

Sub Compliance_Query()


Application.StatusBar = "Creating Queries..."

Dim RPSFromDate As String
RPSFromDate = Format(Range("rps_from_date"), "yyyy-mm-dd hh:mm:ss")
Dim RPSToDate As String
RPSToDate = Format(Range("rps_to_date"), "yyyy-mm-dd hh:mm:ss")

Dim setNC As String
setNC = "SET NOCOUNT ON" & vbNewLine
Dim dropTables As String
dropTables = "DROP TABLE #myPPPseq" & vbNewLine & _
  "DROP TABLE #myCumBudgTotals" & vbNewLine & _
  "DROP TABLE #myBudgTotals" & vbNewLine & _
  "DROP TABLE #myTotals" & vbNewLine & _
  "DROP TABLE #myPiInfo" & vbNewLine & _
  "DROP TABLE #mySupp" & vbNewLine

Dim getTotals As String '-- get postdoc, participant stipend total per project
 getTotals = "SELECT lead, SUM(PDOC_REQ_DOL) AS tot_Pdoc_dol,  SUM(PART_SUPT_STPD_DOL) AS tot_Part_Stipend" & vbNewLine & _
    "INTO #myBudgTotals" & vbNewLine & _
     "FROM (SELECT prop.lead,  budg.PDOC_REQ_DOL, budg.PART_SUPT_STPD_DOL" & vbNewLine & _
        "FROM #myProjPropPI prop, flp.budg budg" & vbNewLine & _
        "WHERE prop.TEMP_PROP_ID = budg.TEMP_PROP_ID ) Tmp" & vbNewLine & _
    "GROUP BY lead" & vbNewLine & _
    "-- get culmulative budget total per project" & vbNewLine & _
    "SELECT lead, revn_num, SUM(budg_tot_dol) AS cumulative_tot_dol" & vbNewLine & _
    "INTO #myCumBudgTotals" & vbNewLine & _
     "FROM (SELECT prop.lead, b.prop_id, b.revn_num, b.budg_tot_dol" & vbNewLine & _
        "FROM #myProjPropPI prop, rptdb.csd.eps_blip b" & vbNewLine & _
       "WHERE prop.PROP_ID = b.PROP_ID And b.revn_num = 0" & vbNewLine & _
        ") Tmp" & vbNewLine & _
    "GROUP BY lead, revn_num" & vbNewLine & _
    "-- get request amount total per project" & vbNewLine & _
    "SELECT lead, SUM(rqst_dol) as tot_rqst_dol, CASE WHEN MIN(prop_titl_txt) <> MAX(prop_titl_txt) THEN 'Y' END AS dif_titl_collab" & vbNewLine & _
    "INTO #myTotals FROM #myProjPropPI GROUP BY lead" & vbNewLine & _
    "-- get other supplement document count for lead proposal only " & vbNewLine & _
    "SELECT prop.lead, count(dtls.supp_doc_seq) as oth_supp_cnt INTO #mySupp" & vbNewLine & _
    "FROM #myProjPropPI prop JOIN flp.supp_dtls dtls ON prop.TEMP_PROP_ID = dtls.TEMP_PROP_ID WHERE prop.ILN < 'M' GROUP BY prop.lead " & vbNewLine

Dim addCOPIs As String '-- reuse the same main table to add all co-pis '- add all co-pis (can mark 'N' to sort all but lead PI together.)
 addCOPIs = "INSERT INTO #myProjPropPI" & vbNewLine & _
    "SELECT 'P' AS ILN, prop.lead, prop.prop_id, prop.TEMP_PROP_ID, prop.nsf_rcvd_date, 0 AS rqst_dol, '' AS prop_titl_txt, addl_pi_invl.pi_id" & vbNewLine & _
    "FROM #myProjPropPI prop" & vbNewLine & _
    "JOIN flp.addl_pi_invl_pars addl_pi_invl ON prop.prop_id = addl_pi_invl.prop_id" & vbNewLine

' This depends on RPSFromDate and RPSToDate
Dim getAwards As String '-- for each PI, use RPS cutoff date to figure out which awards to count
 getAwards = "SELECT PIs.pi_id, awd.awd_id, awd_pi_copi.proj_role_code, awd.awd_eff_date" & vbNewLine & _
    "INTO #myAwds" & vbNewLine & _
    "FROM flp.prop_pars prop, flp.awd_pars awd, flp.awd_pi_copi_pars awd_pi_copi," & vbNewLine & _
    "(SELECT DISTINCT pi_id FROM #myProjPropPI) PIs" & vbNewLine & _
    "WHERE awd_pi_copi.awd_id = awd.awd_id And PIs.pi_id = awd_pi_copi.pi_id" & vbNewLine & _
    "AND awd.awd_id = prop.prop_id AND prop.rcom_awd_istr  not in ('5','8') AND prop.natr_rqst_code not in ('5','A','F') -- remove supplements, etc.;" & vbNewLine & _
    "AND awd.awd_eff_date Between {ts '" & RPSFromDate & "'} And {ts '" & RPSToDate & "'}" & vbNewLine & _
    "-- get per PI awd detail info" & vbNewLine & _
    "SELECT awd.pi_id, count(awd.awd_id) AS NumAwd,  MAX(convert(char(10),awd.awd_eff_date,102)+' '+awd.awd_id+'.'+awd.proj_role_code) AS LastAwd" & vbNewLine & _
    "INTO #myAwdInfo_0" & vbNewLine & _
    "FROM #myAwds awd" & vbNewLine & _
    "GROUP BY awd.pi_id" & vbNewLine & _
    "SELECT b.awd_id, SUM(b.budg_splt_tot_dol) as awd_amt " & vbNewLine & _
    "INTO #myAwdAmt" & vbNewLine & _
    "FROM rptdb.csd.budg_splt b WHERE b.awd_id IN (SELECT DISTINCT SUBSTRING(LastAwd, 12,7) AS LastAwdid FROM #myAwdInfo_0) GROUP BY b.awd_id" & vbNewLine & _
    "SELECT awd.pi_id, awd.NumAwd, awd.LastAwd+' $'+convert(varchar(25),amt.awd_amt) as LastAwd " & vbNewLine & _
    "INTO #myAwdInfo FROM #myAwdInfo_0 awd JOIN #myAwdAmt amt on substring(awd.LastAwd, 12,7) = amt.awd_id " & vbNewLine & _
    "DROP TABLE #myAwdInfo_0  DROP TABLE #myAwdAmt" & vbNewLine & _
    "DROP TABLE #myAwds" & vbNewLine ' parameters here!
  
    
Dim getPerPI As String '-- for each PI, get current & pending support info, which can be text (multiple per pi) as well as pdf upload (1 pdf per pi) ' note that coa and bio are submitted as pdf upload per pi
  getPerPI = "SELECT DISTINCT supt.TEMP_PROP_ID, PIs.pi_id," & vbNewLine & _
    "CASE WHEN CHAR_LENGTH(rtrim(supt.PROJ_TITL_TXT))>0 OR supt.PATH_NAME IS NOT NULL THEN 'Y' ELSE NULL END as cp_chk" & vbNewLine & _
    "INTO #myCPSuppInfo" & vbNewLine & _
    "FROM" & vbNewLine & _
        "(SELECT DISTINCT TEMP_PROP_ID, pi_id FROM #myProjPropPI) PIs" & vbNewLine & _
        "JOIN flp.pi ppi ON PIs.pi_id=ppi.pi_id" & vbNewLine & _
        "JOIN flp.PI_COPI_PRIR_SUPT supt ON PIs.TEMP_PROP_ID=supt.TEMP_PROP_ID AND PIs.pi_id=ppi.pi_id AND ppi.PI_LAST_NAME=supt.PI_LAST_NAME AND ppi.PI_FRST_NAME=supt.PI_FRST_NAME" & vbNewLine & _
    "SELECT id=identity(18), 0 AS seq, p.ILN, p.lead, p.prop_id, p.TEMP_PROP_ID, p.pi_id, pi_copi_char.PI_LAST_NAME, pi_copi_char.pi_emai_addr, inst.inst_shrt_name," & vbNewLine & _
    "a.NumAwd, a.LastAwd, pi_copi_char.PDF_PAGE_CNT AS bio_pg, supt.cp_chk," & vbNewLine & _
    "(SELECT MAX(ca.doc_page_cnt)  FROM flp.clbr_affl ca WHERE p.TEMP_PROP_ID = ca.temp_prop_id AND p.pi_id = ca.pi_id) AS coa_pg" & vbNewLine & _
    "INTO #myPPPseq -- seq: Lead is 0  Non-leads co-Pis" & vbNewLine & _
    "FROM #myProjPropPI p" & vbNewLine & _
    "JOIN flp.pi ppi ON p.pi_id=ppi.pi_id" & vbNewLine & _
    "LEFT OUTER JOIN flp.inst inst ON ppi.inst_id=inst.inst_id -- there some PI has no institution" & vbNewLine & _
    "LEFT OUTER JOIN flp.pi_copi_char pi_copi_char ON p.TEMP_PROP_ID=pi_copi_char.TEMP_PROP_ID AND p.pi_id = pi_copi_char.pi_id" & vbNewLine & _
    "LEFT OUTER JOIN #myAwdInfo a ON p.pi_id = a.pi_id" & vbNewLine & _
    "LEFT OUTER JOIN #myCPSuppInfo supt ON p.TEMP_PROP_ID=supt.TEMP_PROP_ID AND p.pi_id = supt.pi_id" & vbNewLine & _
    "ORDER BY p.lead, p.ILN, a.NumAwd DESC" & vbNewLine
  
Dim cleanUP As String
  cleanUP = "SELECT lead, MIN(id) as 'start'" & vbNewLine & _
            "INTO #myStarts FROM #myPPPseq GROUP BY lead" & vbNewLine & _
    "DROP TABLE #myCPSuppInfo" & vbNewLine & _
    "DROP TABLE #myAwdInfo" & vbNewLine & _
    "DROP TABLE #myProjPropPI" & vbNewLine & _
   "UPDATE #myPPPseq SET p.seq = p.id - M.start" & vbNewLine & _
   "FROM #myPPPseq p, #myStarts M WHERE p.lead = M.lead" & vbNewLine & _
   "DROP TABLE #myStarts" & vbNewLine
  
Dim makePISummary As String
maxPIs = 9
  makePISummary = "SELECT p.lead, MAX(p.seq)+1 AS NumPIs, COUNT(p.bio_pg) AS NumBios,COUNT(p.coa_pg) AS NumCoas,COUNT(p.cp_chk) AS NumCPs, COUNT(p.LastAwd) as NumRPS" & vbNewLine
    For i = 0 To maxPIs
    makePISummary = makePISummary & _
    ", MAX(CASE p.seq WHEN " & i & " THEN p.pi_last_name END) AS pi" & i & "_last_name," & vbNewLine & _
    "MAX(CASE p.seq WHEN " & i & " THEN p.bio_pg END) AS pi" & i & "_bio_pg," & vbNewLine & _
    "max(CASE p.seq WHEN " & i & " THEN p.coa_pg END) AS pi" & i & "_coa_pg," & vbNewLine & _
    "MAX(CASE p.seq WHEN " & i & " THEN p.cp_chk END) AS pi" & i & "_cp_chk," & vbNewLine & _
    "MAX(CASE p.seq WHEN " & i & " THEN p.NumAwd END) AS pi" & i & "_num_awd," & vbNewLine & _
    "MAX(CASE p.seq WHEN " & i & " THEN p.LastAwd END) AS pi" & i & "_last_awd," & vbNewLine & _
    "MAX(CASE p.seq WHEN " & i & " THEN p.pi_emai_addr END) AS pi" & i & "_email," & vbNewLine & _
    "MAX(CASE p.seq WHEN " & i & " THEN p.inst_shrt_name END) AS pi" & i & "_inst" & vbNewLine
   
    Next i


makePISummary = makePISummary & "INTO #myPiInfo FROM #myPPPseq p GROUP BY p.lead" & vbNewLine

  Dim mainQuery As String
  maxPIs = 9
  'Check that su.name (username) is in the list of user names and that they have permission  (103)
  mainQuery = "SELECT CASE when ((SELECT distinct CASE WHEN (su.name in ('public','" & userId & "') and sp.action =193)  then 'True' else 'False' end as Permission FROM FLflpdb.dbo.sysusers su left outer JOIN  FLflpdb.dbo.sysprotects sp  on sp.uid = su.uid left outer JOIN FLflpdb.dbo.sysobjects so ON SU.uid=so.uid where su.name ='" & userId & "'))='False' then 'Need Permission' when (prop.nsf_rcvd_date < '01/25/2016') THEN 'N/A'  when (mpi.NumCoas <1  )then 'Missing' else cast(mpi.NumCoas as varchar) end as NumCoas, " & vbNewLine & _
 "prop.lead_prop_id, prop.prop_id, prop.pm_ibm_logn_id, prop.pgm_annc_id, prop.org_code, prop.pgm_ele_code, " & vbNewLine & _
    "gg.ggpi_trck_num as GG_track_num, gg.ggpi_rcvd_date as GG_rcvd_date, prop.nsf_rcvd_date as NSF_rcvd_date, " & vbNewLine & _
    "CASE WHEN (spcl.SPCL_ITEM_CODE='27') THEN 'Y' END as spcl_excp_flag," & vbNewLine & _
    "prop.prop_stts_code, prop_stts_pars.prop_stts_abbr as prop_stts, ppi.pi_last_name, ppi.pi_frst_name, ppi.pi_emai_addr," & vbNewLine & _
    "INST.inst_shrt_name AS inst_name, INST.st_code, prop.prop_titl_txt, CASE WHEN(CHARINDEX('GOALI',UPPER(prop.prop_titl_txt))>0) THEN 'Y' END AS GOALI_flag," & vbNewLine & _
    "natr_rqst.natr_rqst_txt, prop_stts_pars.prop_stts_txt, prop.cert_dbar_sus_flag, prop_covr.OTH_AGCY_SUBM_FLAG, PROP_COVR.auth_rep_elec_sign, PROP_COVR.AUTH_REP_CERT_DATE," & vbNewLine & _
    "CASE WHEN PROJ_DESC.PDF_PAGE_CNT>0 THEN CAST(PROJ_DESC.PDF_PAGE_CNT AS varchar(4)) ELSE 'No PD' END AS ProjDesc," & vbNewLine & _
    "CASE WHEN (proj_summ.PDF_PAGE_CNT>0 AND proj_summ.SPCL_CHAR_PDF='Y') THEN 'TBC'" & vbNewLine & _
    "WHEN ((proj_summ.PDF_PAGE_CNT<0 OR proj_summ.PDF_PAGE_CNT IS NULL) AND proj_summ.SPCL_CHAR_PDF='Y') THEN 'No PS' END AS ProjSumm," & vbNewLine & _
    "CASE WHEN proj_summ.SPCL_CHAR_PDF='Y' THEN proj_summ.PDF_PAGE_CNT END as 'proj_summ_pg'," & vbNewLine & _
    "(SELECT SUM(BUDG_EXPL.PDF_PAGE_CNT) FROM flp.BUDG_EXPL BUDG_EXPL WHERE c.TEMP_PROP_ID = BUDG_EXPL.TEMP_PROP_ID)  AS 'budg_just_pg'," & vbNewLine & _
    "CASE WHEN (NOT EXISTS ( select 1 from flp.bibl  B where B.TEMP_PROP_ID = c.TEMP_PROP_ID" & vbNewLine & _
        "AND ((B.path_name is not null and  isnull((ltrim(rtrim(path_name))),'') != '') OR ( B.BIBL_TXT NOT LIKE '' AND CHAR_LENGTH(B.BIBL_TXT) IS NOT NULL))))" & vbNewLine & _
        "THEN 'No RC' ELSE CAST(bibl.PDF_PAGE_CNT AS varchar(4)) END as RefCited," & vbNewLine & _
    "CASE WHEN (mbt.tot_Pdoc_dol>0) AND (Not exists ( select 1 from  flp.ment_plan_doc B  where B.TEMP_PROP_ID = c.TEMP_PROP_ID AND B.PATH_NAME IS NOT NULL and  B.PATH_NAME <>  '') )" & vbNewLine & _
        "THEN 'No MP' ELSE CAST(ment.PDF_PAGE_CNT AS varchar(4)) END as MentPlan," & vbNewLine & _
    "CASE WHEN( not exists ( select 1 from flp.dmp_plan_doc B where B.TEMP_PROP_ID = c.TEMP_PROP_ID AND B.PATH_NAME IS NOT NULL and  B.PATH_NAME<>'') )" & vbNewLine & _
        "THEN 'No DMP' ELSE CAST(dmp.PDF_PAGE_CNT AS varchar(4)) END as DMP," & vbNewLine & _
    "CASE WHEN (prop.rcom_awd_istr not in ('5','8','9','C','P','7') and prop.natr_rqst_code not in ('A','5','F','2','X') and prop.obj_clas_code <>'4160'" & vbNewLine & _
         "AND  not exists (SELECT 1 FROM flp.prop_subm_ctl ct, flp.fac_equp a WHERE prop.prop_id = ct.prop_id and ct.temp_prop_id = a.temp_prop_id) )" & vbNewLine & _
         "THEN 'No FAC' END as FacilitiesEquip," & vbNewLine & _
    "mt.dif_titl_collab, mt.tot_rqst_dol, nullif(mbt.tot_Part_Stipend,0) AS tot_Part_Stipend, nullif(mbt.tot_Pdoc_dol,0) AS tot_Pdoc_dol, mcbt.cumulative_tot_dol," & vbNewLine & _
    "nullif(mcbt.cumulative_tot_dol-isnull(mt.tot_rqst_dol,0),0) AS Diff_tot_dol, msupp.oth_supp_cnt, getdate() as run_date, " & vbNewLine & _
    "mpi.*" & vbNewLine

  
Dim mainQuery2 As String
 mainQuery2 = _
 "FROM #myPiInfo mpi" & vbNewLine & _
    "LEFT OUTER JOIN #myBudgTotals mbt ON mpi.lead = mbt.lead" & vbNewLine & _
    "LEFT OUTER JOIN #myTotals mt ON mpi.lead = mt.lead" & vbNewLine & _
    "LEFT OUTER JOIN #myCumBudgTotals mcbt ON mpi.lead = mcbt.lead" & vbNewLine & _
    "LEFT OUTER JOIN #mySupp msupp ON mpi.lead=msupp.lead" & vbNewLine & _
    "JOIN flp.prop_pars prop ON mpi.lead = prop.prop_id" & vbNewLine & _
    "JOIN flp.prop_subm_ctl c ON mpi.lead = c.PROP_ID" & vbNewLine & _
    "JOIN flp.INST INST ON prop.inst_id = INST.inst_id" & vbNewLine & _
    "JOIN flp.pi ppi ON prop.pi_id = ppi.pi_id" & vbNewLine & _
    "LEFT OUTER JOIN flp.prop_stts_pars prop_stts_pars ON prop.prop_stts_code = prop_stts_pars.prop_stts_code" & vbNewLine & _
    "LEFT OUTER JOIN rptdb.csd.natr_rqst natr_rqst ON prop.natr_rqst_code = natr_rqst.natr_rqst_code" & vbNewLine & _
    "LEFT OUTER JOIN flp.PROP_COVR PROP_COVR ON c.TEMP_PROP_ID = PROP_COVR.TEMP_PROP_ID" & vbNewLine & _
    "LEFT OUTER JOIN (select TEMP_PROP_ID, SPCL_ITEM_CODE from flp.PROP_SPCL_ITEM where SPCL_ITEM_CODE ='27') spcl on c.TEMP_PROP_ID=spcl.TEMP_PROP_ID" & vbNewLine & _
    "LEFT OUTER JOIN flp.proj_summ proj_summ ON c.TEMP_PROP_ID = proj_summ.TEMP_PROP_ID" & vbNewLine & _
    "LEFT OUTER JOIN flp.PROJ_DESC PROJ_DESC ON c.TEMP_PROP_ID = PROJ_DESC.TEMP_PROP_ID" & vbNewLine & _
    "LEFT OUTER JOIN flp.bibl bibl ON c.TEMP_PROP_ID = bibl.TEMP_PROP_ID" & vbNewLine & _
    "LEFT OUTER JOIN flp.FAC_EQUP FAC_EQUP ON c.TEMP_PROP_ID = FAC_EQUP.TEMP_PROP_ID" & vbNewLine & _
    "LEFT OUTER JOIN flp.ggpi_trck gg on c.TEMP_PROP_ID=gg.temp_prop_id" & vbNewLine & _
    "LEFT OUTER JOIN flp.MENT_PLAN_DOC ment on c.TEMP_PROP_ID = ment.TEMP_PROP_ID" & vbNewLine & _
    "LEFT OUTER JOIN flp.DMP_PLAN_DOC dmp on c.TEMP_PROP_ID = dmp.TEMP_PROP_ID" & vbNewLine & _
    "ORDER BY prop.lead_prop_id, prop.prop_id " & vbNewLine
   



' ----query for by prop ---
Dim propGetTotal As String
 propGetTotal = "SELECT tp.ILN, prop.lead_prop_id, prop.prop_id, tp.TEMP_PROP_ID, prop.pm_ibm_logn_id, prop.pgm_annc_id,prop.org_code, prop.pgm_ele_code," & vbNewLine & _
    "CASE WHEN prop.org_code <> prop.orig_org_code THEN prop.orig_org_code END AS orig_org_code, " & vbNewLine & _
    "CASE WHEN prop.pgm_ele_code <> prop.orig_pgm_ele_code THEN prop.orig_pgm_ele_code END AS orig_pgm_ele_code," & vbNewLine & _
    "prop.nsf_rcvd_date , prop.rqst_dol, prop.prop_titl_txt, prop.pi_id, prop.prop_stts_code, prop.natr_rqst_code, prop.rcom_awd_istr, prop.obj_clas_code, prop.cert_dbar_sus_flag, prop.inst_id" & vbNewLine & _
    "INTO #myProps" & vbNewLine & _
    "FROM #myProjPropPI tp JOIN flp.prop_pars prop ON tp.prop_id = prop.PROP_ID" & vbNewLine & _
    "SELECT prop1.prop_id, COUNT(prop1.pi_id) AS tot_PIs, SUM(rqst_dol) AS tot_rqst_dol" & vbNewLine & _
    "INTO #myTotals" & vbNewLine & _
    "FROM (SELECT prop.prop_id, prop.pi_id, prop.rqst_dol FROM #myProps prop" & vbNewLine & _
        "Union" & vbNewLine & _
        "SELECT prop.prop_id, addl.pi_id,0 AS rqst_dol FROM flp.addl_pi_invl_pars addl JOIN #myProps prop ON addl.prop_id = prop.prop_id) prop1" & vbNewLine & _
    "GROUP BY prop1.prop_id" & vbNewLine & _
    "SELECT prop.prop_id, Sum(budg.PDOC_REQ_DOL) as tot_Pdoc_dol, Sum(budg.PART_SUPT_STPD_DOL) AS tot_Part_Stipend" & vbNewLine & _
    "INTO #myBudgTotal" & vbNewLine & _
    "FROM flp.budg budg JOIN #myProps prop ON budg.TEMP_PROP_ID = prop.TEMP_PROP_ID" & vbNewLine & _
    "GROUP BY prop.prop_id" & vbNewLine & _
    "SELECT prop.prop_id, count(dtls.supp_doc_seq) as oth_supp_cnt INTO #mySupp" & vbNewLine & _
    "FROM flp.supp_dtls dtls JOIN #myProps prop ON dtls.TEMP_PROP_ID = prop.TEMP_PROP_ID GROUP BY prop.prop_id " & vbNewLine & _
    "SELECT b.prop_id, b.revn_num, SUM(b.budg_tot_dol) AS cumulative_tot_dol" & vbNewLine & _
    "INTO #myCumBudgTotals FROM #myProps prop" & vbNewLine & _
    "JOIN rptdb.csd.eps_blip b ON prop.prop_id = b.PROP_ID" & vbNewLine & _
    "WHERE b.revn_num = 0  GROUP BY b.prop_id, b.revn_num" & vbNewLine

' Use to check permission for time zone
Dim propMainQuery1 As String
 propMainQuery1 = "SELECT prop.ILN, prop.lead_prop_id, prop.prop_id, prop.pm_ibm_logn_id, prop.pgm_annc_id, prop.org_code, prop.pgm_ele_code, " & vbNewLine & _
    "prop.orig_org_code, prop.orig_pgm_ele_code, org1.dir_div_abbr, pe1.pgm_ele_name, org2.dir_div_abbr as orig_dir_div_abbr, pe2.pgm_ele_name as orig_pgm_ele_name, " & vbNewLine & _
    "CASE WHEN (gg.ggpi_rcvd_date is NULL AND INST.last_updt_tmsp>prop.nsf_rcvd_date) THEN INST.last_updt_tmsp " & vbNewLine & _
        "WHEN (gg.ggpi_rcvd_date<>NULL AND INST.last_updt_tmsp>gg.ggpi_rcvd_date) THEN INST.last_updt_tmsp END as inst_last_updt_tmsp," & vbNewLine & _
    "CASE when ((SELECT distinct CASE WHEN (su.name in ('public','" & userId & "') and sp.action =193)  then 'True' else 'False' end FROM FLflpdb.dbo.sysusers su left outer JOIN  FLflpdb.dbo.sysprotects sp  on sp.uid = su.uid left outer JOIN FLflpdb.dbo.sysobjects so ON SU.uid=so.uid where su.name ='" & userId & "'))='False' then 'Need permission: tz_ctry_map' WHEN (gg.ggpi_rcvd_date is NULL AND INST.last_updt_tmsp>prop.nsf_rcvd_date) THEN t.tz_name" & vbNewLine & _
        "WHEN (gg.ggpi_rcvd_date<>NULL AND INST.last_updt_tmsp>gg.ggpi_rcvd_date) THEN t.tz_name END AS inst_timeZone," & vbNewLine & _
    "gg.ggpi_trck_num as GG_track_num, gg.ggpi_rcvd_date as GG_rcvd_date, prop.nsf_rcvd_date as NSF_rcvd_date," & vbNewLine & _
    "CASE WHEN (spcl.SPCL_ITEM_CODE='27') THEN 'Y' END as spcl_excp_flag," & vbNewLine & _
    "prop.prop_stts_code, prop_stts_pars.prop_stts_abbr as prop_stts, ppi.pi_last_name, ppi.pi_frst_name, INST.inst_shrt_name AS inst_name, INST.st_code as inst_st_code, perf_org.perf_org_txt,pi_degr_yr," & vbNewLine & _
    "prop.prop_titl_txt, CASE WHEN(CHARINDEX('GOALI',UPPER(prop.prop_titl_txt))>0) THEN 'Y' END AS GOALI_flag," & vbNewLine & _
    "natr_rqst.natr_rqst_txt, prop_stts_pars.prop_stts_txt," & vbNewLine & _
    "prop.cert_dbar_sus_flag, prop_covr.OTH_AGCY_SUBM_FLAG, PROP_COVR.auth_rep_elec_sign, PROP_COVR.AUTH_REP_CERT_DATE," & vbNewLine & _
    "CASE WHEN PROP_COVR.HUM_DATE is not NULL THEN convert(varchar(10),PROP_COVR.HUM_DATE,1) WHEN PROP_COVR.humn_date_pend_flag='Y' THEN 'Pend' END AS humn_date," & vbNewLine & _
    "CASE WHEN PROP_COVR.VERT_DATE is not NULL THEN convert(varchar(10),PROP_COVR.VERT_DATE,1) WHEN PROP_COVR.vrtb_date_pend_flag='Y' THEN 'Pend' END AS vrtb_date," & vbNewLine

Dim propMainQuery2 As String
 propMainQuery2 = "CASE WHEN ILN='N' THEN 'N/A' WHEN (ILN<>'N' AND PROJ_DESC.PDF_PAGE_CNT>0) THEN CAST(PROJ_DESC.PDF_PAGE_CNT AS varchar(4)) ELSE 'No PD' END AS ProjDesc," & vbNewLine & _
    "CASE WHEN (ILN='N') THEN 'N/A' WHEN (ILN<>'N' AND proj_summ.PDF_PAGE_CNT>0 AND proj_summ.SPCL_CHAR_PDF='Y') THEN 'TBC'" & vbNewLine & _
    "WHEN (ILN<>'N' AND (proj_summ.PDF_PAGE_CNT<0 OR proj_summ.PDF_PAGE_CNT IS NULL) AND proj_summ.SPCL_CHAR_PDF='Y') THEN 'No PS' END AS ProjSumm," & vbNewLine & _
    "CASE WHEN proj_summ.SPCL_CHAR_PDF='Y' THEN proj_summ.PDF_PAGE_CNT END as 'proj_summ_pg',(SELECT SUM(BUDG_EXPL.PDF_PAGE_CNT) FROM flp.BUDG_EXPL BUDG_EXPL WHERE prop.TEMP_PROP_ID = BUDG_EXPL.TEMP_PROP_ID)  AS 'budg_just_pg'," & vbNewLine & _
    "CASE WHEN (NOT EXISTS ( select 1 from flp.bibl  B where B.TEMP_PROP_ID = prop.TEMP_PROP_ID" & vbNewLine & _
        "AND ((B.path_name is not null and  isnull((ltrim(rtrim(path_name))),'') != '') OR ( B.BIBL_TXT NOT LIKE '' AND CHAR_LENGTH(B.BIBL_TXT) IS NOT NULL))))" & vbNewLine & _
        "THEN 'No RC' ELSE CAST(bibl.PDF_PAGE_CNT AS varchar(4)) END as RefCited," & vbNewLine & _
    "CASE WHEN (mb.tot_Pdoc_dol>0) AND (Not exists ( select 1 from  flp.ment_plan_doc B  where B.TEMP_PROP_ID = prop.TEMP_PROP_ID AND B.PATH_NAME IS NOT NULL and  B.PATH_NAME <>  '') )" & vbNewLine & _
        "THEN 'No MP' ELSE CAST(ment.PDF_PAGE_CNT AS varchar(4)) END as MentPlan," & vbNewLine & _
    "CASE WHEN( not exists ( select 1 from flp.dmp_plan_doc B where B.TEMP_PROP_ID = prop.TEMP_PROP_ID AND B.PATH_NAME IS NOT NULL and  B.PATH_NAME<>'') )" & vbNewLine & _
        "THEN 'No DMP' ELSE CAST(dmp.PDF_PAGE_CNT AS varchar(4)) END as DMP," & vbNewLine & _
    "CASE WHEN (prop.rcom_awd_istr not in ('5','8','9','C','P','7') and prop.natr_rqst_code not in ('A','5','F','2','X') and prop.obj_clas_code <>'4160'" & vbNewLine & _
         "AND  not exists (SELECT 1 FROM flp.prop_subm_ctl ct, flp.fac_equp a WHERE prop.prop_id = ct.prop_id and ct.temp_prop_id = a.temp_prop_id) )THEN 'No FAC' END as FacilitiesEquip," & vbNewLine & _
    "mt.tot_PIs, mt.tot_rqst_dol, nullif(mb.tot_Part_Stipend,0) AS tot_Part_Stipend, nullif(mb.tot_Pdoc_dol,0) AS tot_Pdoc_dol, mc.cumulative_tot_dol," & vbNewLine & _
    "nullif(mc.cumulative_tot_dol-isnull(mt.tot_rqst_dol,0),0) as Diff_tot_dol, supp.oth_supp_cnt, getdate() as run_date " & vbNewLine
    
Dim propMainQuery3 As String
propMainQuery3 = _
 "FROM #myProps prop" & vbNewLine & _
    "JOIN flp.pi ppi ON prop.pi_id=ppi.pi_id" & vbNewLine & _
    "JOIN flp.org org1 ON prop.org_code=org1.org_code" & vbNewLine & _
    "JOIN flp.pgm_ele pe1 ON prop.pgm_ele_code=pe1.pgm_ele_code" & vbNewLine & _
    "LEFT OUTER JOIN flp.org org2 ON prop.orig_org_code=org2.org_code" & vbNewLine & _
    "LEFT OUTER JOIN flp.pgm_ele pe2 ON prop.orig_pgm_ele_code=pe2.pgm_ele_code" & vbNewLine & _
    "LEFT OUTER JOIN flp.inst_pars INST ON prop.inst_id = INST.inst_id" & vbNewLine & _
    "LEFT OUTER JOIN rptdb.csd.perf_org perf_org ON perf_org.perf_org_code = INST.perf_org_code" & vbNewLine & _
    "LEFT OUTER JOIN flp.tz_ctry_map t ON inst.inst_tz_id=t.tz_id" & vbNewLine & _
    "LEFT OUTER JOIN flp.prop_stts_pars prop_stts_pars ON prop.prop_stts_code = prop_stts_pars.prop_stts_code" & vbNewLine & _
    "LEFT OUTER JOIN rptdb.csd.natr_rqst natr_rqst ON prop.natr_rqst_code = natr_rqst.natr_rqst_code" & vbNewLine & _
    "LEFT OUTER JOIN flp.PROP_COVR PROP_COVR ON prop.TEMP_PROP_ID = PROP_COVR.TEMP_PROP_ID" & vbNewLine & _
    "LEFT OUTER JOIN (select TEMP_PROP_ID, SPCL_ITEM_CODE from flp.PROP_SPCL_ITEM where SPCL_ITEM_CODE ='27') spcl on prop.TEMP_PROP_ID=spcl.TEMP_PROP_ID" & vbNewLine & _
    "LEFT OUTER JOIN flp.proj_summ proj_summ ON prop.TEMP_PROP_ID = proj_summ.TEMP_PROP_ID" & vbNewLine & _
    "LEFT OUTER JOIN flp.PROJ_DESC PROJ_DESC ON prop.TEMP_PROP_ID = PROJ_DESC.TEMP_PROP_ID" & vbNewLine & _
    "LEFT OUTER JOIN flp.bibl bibl ON prop.TEMP_PROP_ID = bibl.TEMP_PROP_ID" & vbNewLine & _
    "LEFT OUTER JOIN flp.ggpi_trck gg on prop.TEMP_PROP_ID=gg.TEMP_PROP_ID" & vbNewLine & _
    "LEFT OUTER JOIN flp.MENT_PLAN_DOC ment on prop.TEMP_PROP_ID = ment.TEMP_PROP_ID" & vbNewLine & _
    "LEFT OUTER JOIN flp.DMP_PLAN_DOC dmp on prop.TEMP_PROP_ID = dmp.TEMP_PROP_ID" & vbNewLine & _
    "LEFT OUTER JOIN #myTotals mt ON prop.prop_id = mt.prop_id" & vbNewLine & _
    "LEFT OUTER JOIN #myBudgTotal mb ON prop.prop_id= mb.prop_id" & vbNewLine & _
    "LEFT OUTER JOIN #myCumBudgTotals mc ON prop.prop_id = mc.prop_id" & vbNewLine & _
    "LEFT OUTER JOIN #mySupp supp ON prop.prop_id=supp.prop_id" & vbNewLine & _
    "ORDER BY prop.lead_prop_id, prop.ILN, prop.prop_id" & vbNewLine

Dim propDropTables As String
 propDropTables = "DROP TABLE #myProps" & vbNewLine & _
    "DROP TABLE #myTotals" & vbNewLine & _
    "DROP TABLE #myBudgTotal" & vbNewLine & _
    "DROP TABLE #myCumBudgTotals" & vbNewLine & _
    "DROP TABLE #myProjPropPI" & vbNewLine & _
    "DROP TABLE #mySupp" & vbNewLine


' MsgBox setNC & selectProps & mainQuery & dropProps
'ODBC;dsn=RPTSERVER;database=rptdb;server=RPTSQL.nsf.gov;port=5000;UID=" & userId & ";PWD=" & pwd & ";", Destination:=Range("$A$1")).QueryTable
    
    
    Dim QT2 As QueryTable
    Set QT2 = Worksheets("Compliance-ByProp").ListObjects.Item(1).QueryTable
    With QT2
        .CommandText = setNC & selectProps & propGetTotal & propMainQuery1 & propMainQuery2 & propMainQuery3 & propDropTables
        
        .Refresh (False)
    End With
    
Application.StatusBar = "done ByProp"
    
    Dim QT As QueryTable
    Set QT = Worksheets("Compliance-ByProj").ListObjects.Item(1).QueryTable
    With QT
        .CommandText = setNC & selectProps & getTotals & addCOPIs & getAwards & getPerPI & cleanUP & _
        makePISummary & mainQuery & mainQuery2 & dropTables
        .Refresh BackgroundQuery:=False
    End With
    
Application.StatusBar = "done ByProj"
    
    Dim pivotTable As pivotTable
    For Each pivotTable In Worksheets("Dashboard").PivotTables
       On Error Resume Next
       pivotTable.RefreshTable
       On Error GoTo 0
    Next
    
Application.StatusBar = False
  
End Sub

