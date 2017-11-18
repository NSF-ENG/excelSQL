SET NOCOUNT ON 
SELECT CASE WHEN p.lead_prop_id IS NULL THEN 'I' WHEN p.lead_prop_id <> p.prop_id THEN 'N' ELSE 'L' END AS ILN,
isnull(p.lead_prop_id,p.prop_id) AS lead, p.prop_id, psc.TEMP_PROP_ID
INTO #myPids FROM csd.prop p
JOIN csd.prop_subm_ctl_vw psc ON psc.prop_id = p.prop_id
JOIN csd.panl_prop pp ON p.prop_id = pp.prop_id
WHERE --p.prop_stts_code IN ('00','01','02','08','09') AND
 pp.panl_id in ('p180207')
ORDER BY lead, ILN, p.prop_id
CREATE INDEX myPid_idx ON #myPids(prop_id)
-- we should have everyone already, but let's make sure
INSERT INTO #myPids SELECT CASE WHEN pr.lead_prop_id <> pr.prop_id THEN 'N' ELSE 'L' END as ILN, 
p.lead, p.prop_id, psc.TEMP_PROP_ID
FROM #myPids p 
JOIN csd.prop pr ON pr.lead_prop_id = p.lead
JOIN csd.prop_subm_ctl_vw psc ON psc.prop_id = pr.prop_id
WHERE p.ILN <> 'I' 
AND NOT EXISTS (SELECT * FROM #myPids px WHERE px.prop_id = pr.prop_id) -- skip what we have already
ORDER BY lead, ILN, p.prop_id
CREATE INDEX myPid_idx3 ON #myPids(ILN, lead, prop_id) 
CREATE INDEX myPid_temp ON #myPids(TEMP_PROP_ID)
SELECT lead, convert(text, convert(varchar(16384),js.PROJ_SUMM_TXT) + convert(varchar(16384),js.INTUL_MERT) + convert(varchar(16384),js.BRODR_IMPT)) as summ
INTO #mySumm
FROM #myPids p 
JOIN FLflpdb.flp.proj_summ js ON js.SPCL_CHAR_PDF <> 'Y' AND js.TEMP_PROP_ID = p.TEMP_PROP_ID
WHERE p.ILN < 'M'
SELECT prop.nsf_rcvd_date, nullif(prop.dd_rcom_date,'1900-01-01') AS dd_rcom_date,
prop.pgm_annc_id, o2.dir_div_abbr as Dir, prop.org_code, CASE WHEN prop.org_code <> prop.orig_org_code THEN prop.orig_org_code END AS origORG, 
prop.pgm_ele_code+' - '+pgm_ele_name as Pgm, CASE WHEN prop.pgm_ele_code <> prop.orig_pgm_ele_code THEN prop.orig_pgm_ele_code ELSE ' ' END AS origPEC,
prop.pm_ibm_logn_id as PO, prop.obj_clas_code, natr_rqst.natr_rqst_abbr, prop_stts.prop_stts_abbr, p.ILN, p.lead, org.dir_div_abbr as Div, p.prop_id,p.TEMP_PROP_ID,
pi.pi_last_name, pi.pi_frst_name, inst.inst_shrt_name AS inst_name, inst.st_code, pi.pi_emai_addr,
prop.prop_titl_txt, prop.rqst_dol, prop.rqst_eff_date, prop.rqst_mnth_cnt, prop.cntx_stmt_id, prop.prop_stts_code, prop_stts.prop_stts_txt, prop.inst_id, prop.pi_id
INTO #myProps FROM (SELECT DISTINCT * FROM #myPids ) p
JOIN csd.prop prop ON p.prop_id = prop.prop_id
JOIN csd.org org ON prop.org_code = org.org_code
JOIN csd.org o2 ON left(prop.org_code,2)+'000000' = o2.org_code
JOIN csd.pgm_ele pe ON prop.pgm_ele_code = pe.pgm_ele_code
JOIN csd.inst inst ON prop.inst_id = inst.inst_id
JOIN csd.pi_vw pi ON prop.pi_id = pi.pi_id
JOIN csd.prop_stts prop_stts ON prop.prop_stts_code = prop_stts.prop_stts_code
JOIN csd.natr_rqst natr_rqst ON natr_rqst.natr_rqst_code = prop.natr_rqst_code
WHERE pi.prim_addr_flag='Y'
ORDER BY lead, ILN DROP TABLE #myPids CREATE INDEX p1 ON #myProps(lead)
SELECT p.lead, p.pi_last_name, rs.string, rs.score, rev_prop_vw.revr_id, rev_prop_vw.rev_rtrn_date, id=identity(18), 0 as 'seq'
INTO #myRevs
FROM #myProps p, csd.rev_prop_vw rev_prop_vw, tempdb.guest.revScores rs
WHERE p.ILN < 'M' AND p.lead = rev_prop_vw.prop_id AND rev_prop_vw.rev_prop_rtng_ind = rs.yn 
ORDER BY lead, score DESC

SELECT lead, MIN(id) as 'start' INTO #myStarts FROM #myRevs GROUP BY lead
UPDATE #myRevs set seq = id-M.start FROM #myRevs r, #myStarts M WHERE r.lead = M.lead
DROP TABLE #myStarts
SELECT   rv.lead, rv.pi_last_name, 'Review' as docType, rev_prop.pm_logn_id, rv.revr_id as panl_revr_id,  revr.revr_last_name as 'name', revr_opt_addr_line.revr_addr_txt as 'info',
convert(varchar,rev_prop.rev_type_code) AS type, convert(varchar,rev_prop.rev_stts_code) as stts, rev_prop.rev_due_date as due, rv.rev_rtrn_date as returned,
 rv.string as score, rev_txt.REV_PROP_TXT_FLDS as 'text'
FROM #myRevs rv, csd.revr revr, csd.rev_prop rev_prop, csd.revr_opt_addr_line revr_opt_addr_line, csd.rev_prop_txt_flds_vw rev_txt
WHERE rv.lead = rev_prop.prop_id AND rv.lead = rev_txt.PROP_ID AND rv.revr_id = rev_prop.revr_id AND rv.revr_id = revr.revr_id AND rv.revr_id = revr_opt_addr_line.revr_id
  AND rv.revr_id = rev_txt.REVR_ID AND ((revr_opt_addr_line.addr_lne_type_code='E'))
UNION ALL SELECT p.lead, p.pi_last_name, ' PanlSumm', panl.pm_logn_id, panl_prop_summ.PANL_ID,  ' '+panl.panl_name, panl_rcom_def.RCOM_TXT,
convert(varchar,panl_prop_summ.RCOM_SEQ_NUM), convert(varchar,panl_prop_summ.PROP_ORDR), panl.panl_bgn_date, panl_prop_summ.panl_summ_rlse_date,
panl_rcom_def.RCOM_ABBR , panl_prop_summ.PANL_SUMM_TXT
FROM #myProps p, FLflpdb.flp.panl_prop_summ panl_prop_summ, FLflpdb.flp.panl_rcom_def panl_rcom_def, csd.panl panl
WHERE  p.ILN < 'M' AND p.lead = panl_prop_summ.PROP_ID AND panl_prop_summ.PANL_ID = panl.panl_id
  AND panl_prop_summ.RCOM_SEQ_NUM = panl_rcom_def.RCOM_SEQ_NUM AND panl_prop_summ.PANL_ID = panl_rcom_def.PANL_ID
UNION ALL SELECT p.lead, p.pi_last_name, 'POCmnt', p.PO, p.prop_id, cmt.cmnt_cre_id, '', '', convert(varchar,cmt.cmnt_prop_stts_code), cmt.beg_eff_date, cmt.end_eff_date,'', cmt.cmnt
FROM #myProps p, FLflpdb.flp.cmnt_prop cmt WHERE p.prop_id = cmt.prop_id AND (p.ILN < 'M' OR LEN(cmt.cmnt) <> (SELECT LEN(l.cmnt) FROM FLflpdb.flp.cmnt_prop l WHERE p.lead = l.prop_id))
UNION ALL SELECT p.lead, p.pi_last_name, 'RA' as docType, p.PO, '', ra.last_updt_user, '', '', null, null, ra.last_updt_tmsp, '', ra.prop_rev_anly_txt
FROM #myProps p, csd.prop_rev_anly_vw ra WHERE p.ILN < 'M' AND p.lead = ra.prop_id
UNION ALL SELECT p.lead, p.pi_last_name, 'Abstr', p.PO, '', a.last_updt_user, a.cent_mrkr_prop, a.cent_mrkr_awd, null, null, a.last_updt_tmsp,'', a.abst_narr_txt
FROM #myProps p, csd.abst a WHERE p.prop_id = a.awd_id AND (p.ILN < 'M' OR LEN(a.abst_narr_txt) <>  (SELECT LEN(l.abst_narr_txt) FROM csd.abst l WHERE p.lead = l.awd_id))
UNION ALL SELECT p.lead, p.pi_last_name, 'SummProj',p.PO, '', '', '', null, null, null, p.nsf_rcvd_date,'', s.summ
FROM #myProps p, #mySumm s WHERE p.ILN < 'M' AND s.lead = p.lead
UNION ALL SELECT p.lead, p.pi_last_name, 'xDiaryNt',p.PO, p.prop_id, crtd_by_user, ej_diry_note_kywd, null, null, null, crtd_date,'', ej_diry_note_txt
FROM #myProps p, FLflpdb.flp.ej_diry_note d WHERE d.prop_id = p.prop_id
ORDER BY lead, docType, revr_last_name, revr_id
DROP TABLE #myRevs,#myProps, #mySumm

