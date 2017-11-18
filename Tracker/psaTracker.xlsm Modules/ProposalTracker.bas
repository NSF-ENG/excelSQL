Attribute VB_Name = "ProposalTracker"
Sub ClearPropInputs()
' Clear inputs
    ActiveSheet.Range("B5:C16").Cells.Value = Settings.Range("A20:B31").Cells.Value
    Call ClearTable(ActiveSheet.ListObjects("add_prop_ids"))
    Call ClearTable(ActiveSheet.ListObjects("omit_prop_ids"))
End Sub

Sub RefreshActiveSheetProp()
Dim field As String
Dim dateWhere As String
Dim dateDDWhere As String 'dd_rcm_date range
Dim addProps As String
Dim start As String ' True or false to start where

Call ShowPwdForm

addProps = IDsFromColumnRange("OR prop.prop_id IN", "add_prop_ids")

If hasValue("from_date") Then dateWhere = "AND prop.nsf_rcvd_date >= {ts '" & Format(ActiveSheet.Range("from_date"), "yyyy-mm-dd hh:mm:ss") & "'}"
If hasValue("to_date") Then dateWhere = dateWhere & " AND prop.nsf_rcvd_date <= {ts '" & Format(ActiveSheet.Range("to_date"), "yyyy-mm-dd hh:mm:ss") & "'} "
If hasValue("dd_from_date") Then dateDDWhere = "AND prop.dd_rcom_date >= {ts '" & Format(ActiveSheet.Range("dd_from_date"), "yyyy-mm-dd hh:mm:ss") & "'}"
If hasValue("dd_to_date") Then dateDDWhere = dateDDWhere & " AND prop.dd_rcom_date <= {ts '" & Format(ActiveSheet.Range("dd_to_date"), "yyyy-mm-dd hh:mm:ss") & "'} "

If (hasValue("from_date") Or hasValue("dd_from_date")) Then
     start = "(1=1)" 'True
Else
    start = "(0=1)" 'False
    If Len("add_prop_ids") < 1 Then ' no specific prop_id
        MsgBox "Include both dates, or include proposal numbers in the Add table."
        End
    End If
End If

'-----------  tracking with temp tables
Dim setNC As String
setNC = "set nocount on" & vbNewLine
Dim myProp As String
myProp = "SELECT isnull(prop.lead_prop_id,prop.prop_id) AS 'lead_id'," & vbNewLine _
& "CASE WHEN prop.lead_prop_id IS NULL THEN 'I' WHEN prop.lead_prop_id = prop.prop_id THEN 'L' ELSE 'N' END AS ILN, prop.prop_id," & vbNewLine _
& "(SELECT MAX( CASE pa.prop_atr_seq WHEN 1 THEN pa.prop_atr_code ELSE '' END ) + ' ' +" & vbNewLine _
& "        MAX( CASE pa.prop_atr_seq WHEN 2 THEN pa.prop_atr_code ELSE '' END ) + ' ' +" & vbNewLine _
& "        MAX( CASE pa.prop_atr_seq WHEN 3 THEN pa.prop_atr_code ELSE '' END ) + ' ' +" & vbNewLine _
& "        MAX( CASE pa.prop_atr_seq WHEN 4 THEN pa.prop_atr_code ELSE '' END ) + ' ' +" & vbNewLine _
& "        MAX( CASE pa.prop_atr_seq WHEN 5 THEN pa.prop_atr_code ELSE '' END ) + ' ' +" & vbNewLine _
& "        MAX( CASE pa.prop_atr_seq WHEN 6 THEN pa.prop_atr_code ELSE '' END )" & vbNewLine _
& "        FROM csd.prop_atr pa  WHERE pa.prop_id = prop.prop_id  AND pa.prop_atr_type_code = 'PRC' ) AS 'PRCs'," & vbNewLine
myProp = myProp _
& "(SELECT Count(*) FROM csd.rev_prop rp WHERE prop.prop_id=rp.prop_id AND rp.rev_type_code='R')  AS 'adhoc_reqd'," & vbNewLine _
& "(SELECT Count(*) FROM csd.rev_prop rp WHERE prop.prop_id=rp.prop_id AND rp.rev_type_code='R' AND rp.rev_stts_code='P')  AS 'adhoc_pend'," & vbNewLine _
& "(SELECT Max(rp.rev_due_date) FROM csd.rev_prop rp WHERE prop.prop_id=rp.prop_id AND rp.rev_type_code='R' AND rp.rev_stts_code='P')  AS 'last_adhoc_due'," & vbNewLine _
& "(SELECT Count(*) FROM csd.rev_prop rp WHERE prop.prop_id=rp.prop_id AND rp.rev_stts_code='R')  AS 'revRcvd'," & vbNewLine _
& "(SELECT Max(rp.rev_due_date) FROM csd.rev_prop rp WHERE prop.prop_id=rp.prop_id)  AS 'last_rev_due', prop.nsf_rcvd_date,nullif(prop.dd_rcom_date,'1/1/1900') AS dd_rcom_date" & vbNewLine _
& "INTO #myProp" & vbNewLine _
& "FROM csd.prop prop" & vbNewLine _
& "JOIN csd.prop_stts ps on ps.prop_stts_code=prop.prop_stts_code" & vbNewLine _
& "JOIN csd.natr_rqst nr on nr.natr_rqst_code = prop.natr_rqst_code" & vbNewLine _
& "JOIN csd.org  as og on og.org_code=prop.org_code" & vbNewLine _
& "WHERE ((" & start & dateWhere & dateDDWhere & vbNewLine _
& andWhere("prop", "pgm_annc_id") & andWhere("prop", "org_code") & vbNewLine _
& andWhere("prop", "pgm_ele_code") & andWhere("prop", "pm_ibm_logn_id") & vbNewLine _
& andWhere("ps", "prop_stts_abbr") & andWhere("prop", "obj_clas_code") & vbNewLine _
& andWhere("nr", "natr_rqst_abbr") & vbNewLine _
& andWhere("og", "dir_div_abbr") & vbNewLine _
& andWhere("prop", "prop_titl_txt") & vbNewLine _

'-----------CASE FOR PROP PRCS INCLUDE/EXCLUDE--------------------------------------
field = Trim(ActiveSheet.Range("pa.prop_atr_code").Value)
If Left(field, 1) = "~" Then  ' have negation,Prop PRCS.
myProp = myProp _
    & excludePRCS("csd.prop_atr pa ", "", "pa.prop_atr_code", " and pa.prop_id=prop.prop_id AND pa.prop_atr_type_code='PRC' ") & vbNewLine
Else ' Include Budg PRCS

myProp = myProp _
    & includePRCS("csd.prop_atr pa ", "", "pa.prop_atr_code", " and pa.prop_id=prop.prop_id AND pa.prop_atr_type_code='PRC' ") & vbNewLine
End If

myProp = myProp _
& ") " & addProps & ") " & IDsFromColumnRange("AND prop.prop_id NOT IN", "omit_prop_ids") & vbNewLine _
& "ORDER BY lead_id,ILN" & vbNewLine

Dim myPanl As String
myPanl = "SELECT panl_prop.prop_id, panl_prop.panl_id, panl.panl_bgn_date, a.rcom_seq_num, b.rcom_abbr, a.prop_ordr" & vbNewLine _
& "INTO #myPanl" & vbNewLine _
& "FROM #myProp prop, csd.panl_prop panl_prop, csd.panl panl, flflpdb.flp.panl_prop_summ a, flflpdb.flp.panl_rcom_def b" & vbNewLine _
& "WHERE  prop.prop_id=panl_prop.prop_id AND panl_prop.panl_id = panl.panl_id" & vbNewLine _
& "AND  panl_prop.panl_id *= a.panl_id AND prop.prop_id *= a.prop_id AND a.panl_id *= b.panl_id  AND  a.rcom_seq_num *= b.rcom_seq_num" & vbNewLine '---- allow missing recom

Dim Query As String
Query = "SELECT getdate() as run_date,mp.*, prop.pgm_annc_id, prop.org_code, prop.pgm_ele_code, prop.pm_ibm_logn_id, prop_stts.prop_stts_abbr, prop.prop_stts_code, prop_stts.prop_stts_txt, pi.pi_last_name, pi.pi_frst_name, pi.pi_gend_code, inst.inst_shrt_name AS inst_name, inst.st_code, prop.prop_titl_txt, natr_rqst.natr_rqst_txt, natr_rqst.natr_rqst_abbr, prop.bas_rsch_pct, prop.cntx_stmt_id," & vbNewLine _
& "first.panl_id as 'first_panl', first.panl_bgn_date as 'fp_begin', first.rcom_seq_num as 'fp_recno', first.rcom_abbr as 'fp_rec', first.prop_ordr as 'fp_rank', last.panl_id as 'last_panl', last.panl_bgn_date as 'lp_begin', last.rcom_seq_num as 'lp_recno', last.rcom_abbr as 'lp_rec', last.prop_ordr as 'lp_rank'," & vbNewLine _
& "bs.split_tot_dol, bs.split_frwd_date, bs.split_aprv_date," & vbNewLine _
& "prop.rqst_dol, prop.rqst_mnth_cnt, nullif(prop.rcom_mnth_cnt,0) AS 'rcom_mnth_cnt', prop.rqst_eff_date, nullif(prop.rcom_eff_date,'1900-01-01') AS 'rcom_eff_date', nullif(prop.pm_asgn_date,'1900-01-01') AS pm_asgn_date, nullif(prop.pm_rcom_date,'1900-01-01') AS  pm_rcom_date, nullif(prop.dd_rcom_date,'1900-01-01') AS  dd_rcom_date," & vbNewLine

Query = Query & "awd.awd_id, awd.tot_intn_awd_amt, pi2.pi_last_name, pi2.pi_frst_name, inst2.inst_shrt_name AS inst_awd, awd.awd_titl_txt, awd.pm_ibm_logn_id, awd.org_code, awd.pgm_ele_code, awd.pgm_div_code, awd.awd_istr_code, awd.awd_stts_code, awd.fpr_stts_code, awd.awd_stts_date, awd.awd_eff_date, awd.awd_exp_date, awd.awd_fin_clos_date, awd.fpr_stts_updt_date, awd.est_fnl_exp_date" & vbNewLine _
& "FROM  #myProp mp, csd.prop prop, csd.inst inst, csd.natr_rqst natr_rqst, csd.pi pi, csd.prop_stts prop_stts,  csd.awd awd, csd.inst inst2, csd.pi pi2," & vbNewLine _
& "(SELECT *  FROM #myPanl pn" & vbNewLine _
    & "WHERE  pn.panl_bgn_date =(SELECT min(p.panl_bgn_date) FROM #myPanl p  WHERE pn.prop_id=p.prop_id ) ) first," & vbNewLine _
& "(SELECT *  FROM #myPanl pn" & vbNewLine _
    & "WHERE  pn.panl_bgn_date >(SELECT min(p.panl_bgn_date) FROM #myPanl p WHERE pn.prop_id=p.prop_id )" & vbNewLine _
        & "AND pn.panl_bgn_date = (SELECT max(p.panl_bgn_date) FROM #myPanl p  WHERE pn.prop_id=p.prop_id ) ) last," & vbNewLine
        
Query = Query & "(SELECT budg_splt.prop_id, Sum(budg_splt.budg_splt_tot_dol) AS 'split_tot_dol', Max(budg_splt.frwd_date) AS 'split_frwd_date', Max(budg_splt.aprv_date) AS 'split_aprv_date'" & vbNewLine _
& "FROM csd.budg_splt budg_splt GROUP BY budg_splt.prop_id) bs" & vbNewLine _
& "WHERE mp.prop_id = prop.prop_id AND prop.natr_rqst_code = natr_rqst.natr_rqst_code AND prop.prop_stts_code = prop_stts.prop_stts_code AND prop.inst_id = inst.inst_id AND prop.pi_id = pi.pi_id AND" & vbNewLine _
& "mp.prop_id *= first.prop_id AND mp.prop_id *= last.prop_id AND mp.prop_id *= bs.prop_id AND mp.prop_id *= awd.awd_id AND awd.inst_id *= inst2.inst_id AND awd.pi_id *= pi2.pi_id" & vbNewLine _
& "ORDER BY mp.lead_id, mp.ILN, mp.prop_id" & vbNewLine

Dim dropTables As String
dropTables = "drop table #myProp drop table #myPanl" & vbNewLine

    Dim QT As QueryTable
    Dim LO As ListObject
    For Each LO In ActiveSheet.ListObjects
      If (Left(LO.Name, 14) = "PropQueryTable") Then Set QT = LO.QueryTable 'excel adds a number; we ignore.
    Next
    With QT
     .CommandText = setNC & myProp & myPanl & Query & dropTables
     .Refresh (True)
    End With
End Sub




