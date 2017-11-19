-- PD-3PO v1.1 sql code

SET NOCOUNT ON
SELECT CASE WHEN prop.lead_prop_id IS NULL THEN 'I' WHEN prop.lead_prop_id <> prop.prop_id THEN 'N' ELSE 'L' END AS ILN,
isnull(prop.lead_prop_id,prop.prop_id) AS lead, prop.prop_id
INTO #myPids FROM csd.prop prop
JOIN csd.prop_stts prop_stts ON prop.prop_stts_code = prop_stts.prop_stts_code
JOIN csd.natr_rqst natr_rqst ON prop.natr_rqst_code = natr_rqst.natr_rqst_code
JOIN csd.org org ON prop.org_code = org.org_code
WHERE prop.dd_rcom_date >= {ts '2015-10-01 00:00:00'} 
AND prop.dd_rcom_date <= {ts '2016-12-09 00:00:00'} 
 AND prop.pgm_ele_code IN ('5373','1591')
 AND prop_stts.prop_stts_abbr = 'AWD'
 AND natr_rqst.natr_rqst_abbr = 'NEW'
 AND org.dir_div_abbr = 'IIP'
INSERT INTO #myPids SELECT CASE WHEN prop.lead_prop_id <> prop.prop_id THEN 'N' ELSE 'L' END as ILN, p.lead, prop.prop_id
FROM #myPids p, csd.prop prop WHERE p.ILN <> 'I' AND p.lead = prop.lead_prop_id
SELECT prop.nsf_rcvd_date, nullif(prop.dd_rcom_date,'1900-01-01') AS dd_rcom_date,
prop.pgm_annc_id, o2.dir_div_abbr as Dir, prop.org_code, CASE WHEN prop.org_code <> prop.orig_org_code THEN prop.orig_org_code END AS origORG, 
prop.pgm_ele_code+' - '+pgm_ele_name as Pgm, CASE WHEN prop.pgm_ele_code <> prop.orig_pgm_ele_code THEN prop.orig_pgm_ele_code ELSE ' ' END AS origPEC,
prop.pm_ibm_logn_id as PO, prop.obj_clas_code, natr_rqst.natr_rqst_abbr, prop_stts.prop_stts_abbr, p.ILN, p.lead, org.dir_div_abbr as Div, p.prop_id,
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
ORDER BY lead, ILN DROP TABLE #myPids

create table #revScores(yn char(5) primary key, string varchar(10), score real null)
insert into #revScores
select 'NNNNN', 'R', null  union all
select 'NNNNY', 'P', 1 union all
select 'NNNYN', 'F', 3 union all
select 'NNNYY', 'F/P', 2 union all
select 'NNYNN', 'G', 5 union all
select 'NNYNY', 'G/P', 2.98 union all
select 'NNYYN', 'G/F', 4 union all
select 'NNYYY', 'G/F/P', 2.99 union all
select 'NYNNN', 'V', 7 union all
select 'NYNNY', 'V/P', 3.98 union all
select 'NYNYN', 'V/F', 4.98 union all
select 'NYNYY', 'V/F/P', 3.65 union all
select 'NYYNN', 'V/G', 6 union all
select 'NYYNY', 'V/G/P', 4.32 union all
select 'NYYYN', 'V/G/F', 4.99 union all
select 'NYYYY', 'V/G/F/P', 3.97 union all
select 'YNNNN', 'E', 9 union all
select 'YNNNY', 'E/P', 4.992 union all
select 'YNNYN', 'E/F', 5.98 union all
select 'YNNYY', 'E/F/P', 4.325 union all
select 'YNYNN', 'E/G', 6.98 union all
select 'YNYNY', 'E/G/P', 4.995 union all
select 'YNYYN', 'E/G/F', 5.66 union all
select 'YNYYY', 'E/G/F/P', 4.5 union all
select 'YYNNN', 'E/V', 8 union all
select 'YYNNY', 'E/V/P', 5.666 union all
select 'YYNYN', 'E/V/F', 6.33 union all
select 'YYNYY', 'E/V/F/P', 4.996 union all
select 'YYYNN', 'E/V/G', 6.99 union all
select 'YYYNY', 'E/V/G/P', 5.5 union all
select 'YYYYN', 'E/V/G/F', 5.99 union all
select 'YYYYY', 'E/V/G/F/P', 4.997



-- Get personel from prop, addl_pi_invl, sr_pers, subaward
-- Assumes #myProps is defined.
-- todo: add biosketch, coa, curr&pend

--[PD3_allPIs
SELECT prop.prop_id, prop.ILN, prop.prop_titl_txt, ' PI' as pers, pi.pi_last_name, pi.pi_frst_name, prop.inst_name,  pi.pi_gend_code, pi.pi_emai_addr, pi.pi_dept_name, prop.st_code, prop.inst_id, pi.pi_id, prop.rqst_dol
INTO #myPers 
FROM #myProps prop, csd.pi_vw pi
WHERE prop.pi_id = pi.pi_id AND (pi.prim_addr_flag='Y') 
UNION ALL SELECT prop.prop_id, prop.ILN, '', 'coPI', pi.pi_last_name, pi.pi_frst_name, inst.inst_shrt_name,  pi.pi_gend_code, pi.pi_emai_addr, pi.pi_dept_name, pi.st_code, pi.inst_id, pi.pi_id, 0
FROM csd.addl_pi_invl addl_pi_invl, csd.inst inst, csd.pi_vw pi, #myProps prop
  WHERE prop.prop_id = addl_pi_invl.prop_id AND addl_pi_invl.pi_id = pi.pi_id AND pi.inst_id = inst.inst_id AND (pi.prim_addr_flag='Y') 
CREATE INDEX myPersPIx  ON #myPers(prop_id)
CREATE INDEX myPersPIy ON #myPers(pi_last_name, pi_frst_name, inst_name)

INSERT #myPers (prop_id, ILN, prop_titl_txt, pers, pi_last_name, pi_frst_name, inst_name, pi_gend_code, pi_emai_addr, pi_dept_name, st_code, inst_id, pi_id, rqst_dol)
SELECT DISTINCT prop.prop_id, prop.ILN, '', 'srPers', s.SR_LAST_NAME AS pi_last_name, s.SR_FRST_NAME AS pi_frst_name, inst.inst_shrt_name AS inst_name, '', '', '', inst.st_code, inst.inst_id, '', 0
FROM #myProps prop
JOIN csd.prop_subm_ctl_vw psc ON prop.prop_id = psc.PROP_ID
JOIN csd.sr_pers_resc_vw s ON psc.TEMP_PROP_ID = s.TEMP_PROP_ID
JOIN csd.inst inst ON s.PERF_INST_ID = inst.inst_id
  WHERE NOT EXISTS (SELECT * FROM #myPers p
WHERE prop.prop_id = p.prop_id  AND s.SR_LAST_NAME = p.pi_last_name AND s.SR_FRST_NAME = pi_frst_name AND inst.inst_shrt_name = p.inst_name)
INSERT #myPers (prop_id, ILN, prop_titl_txt, pers, pi_last_name, pi_frst_name, inst_name, pi_gend_code, pi_emai_addr, pi_dept_name, st_code, inst_id, pi_id, rqst_dol)
SELECT DISTINCT prop.prop_id, prop.ILN, '', 'subAwd', s.PI_LAST_NAME AS pi_last_name, s.PI_FRST_NAME AS pi_frst_name, inst.inst_shrt_name AS inst_name, '', '', '', inst.st_code, inst.inst_id, '', 0
FROM #myProps prop
JOIN csd.prop_subm_ctl_vw psc ON prop.prop_id = psc.PROP_ID
JOIN csd.subconpi_vw s ON psc.TEMP_PROP_ID = s.TEMP_PROP_ID
JOIN csd.inst inst ON s.PERF_INST_ID = inst.inst_id
 WHERE NOT EXISTS (SELECT * FROM #myPers p
   WHERE prop.prop_id = p.prop_id  AND s.PI_LAST_NAME = p.pi_last_name AND s.PI_FRST_NAME = pi_frst_name AND inst.inst_shrt_name = p.inst_name)SELECT  prop.nsf_rcvd_date, prop.dd_rcom_date, prop.pgm_annc_id, prop.Dir, prop.org_code, prop.Pgm, prop.PO,
 prop.obj_clas_code, prop.natr_rqst_abbr, prop.prop_stts_abbr, prop.lead, prop.Div, pr.*, prop.rqst_mnth_cnt, prop.rqst_eff_date
FROM #myProps prop, #myPers pr
WHERE pr.prop_id = prop.prop_id
ORDER BY prop.lead, pr.ILN, pr.prop_id, pr.pers, pr.pi_last_name

DROP TABLE #myPers  DROP TABLE #myProps
--]PD3_allPIs

--[PD3_props
SELECT sp1.TEMP_PROP_ID, ctry.ctry_name, id=identity(18), 0 as 'seq' INTO #myCtry
FROM #myProps prop
JOIN csd.prop_subm_ctl_vw psc ON prop.prop_id = psc.prop_id
JOIN csd.prop_spcl_item_vw sp1 ON sp1.TEMP_PROP_ID = psc.TEMP_PROP_ID
JOIN csd.ctry ctry ON sp1.SPCL_ITEM_CODE = ctry.ctry_code
WHERE end_date Is Null
ORDER BY sp1.TEMP_PROP_ID, ctry.ctry_name
SELECT TEMP_PROP_ID, MIN(id) as 'start' INTO #myStart FROM #myCtry GROUP BY TEMP_PROP_ID
UPDATE #myCtry set seq = id-M.start FROM #myCtry r, #myStart M WHERE r.TEMP_PROP_ID = M.TEMP_PROP_ID
DROP TABLE #myStart
SELECT DISTINCT prop.prop_id, pa.prop_atr_code, id=identity(18), 0 as 'seq' INTO #myPRCs
FROM #myProps prop, csd.prop_atr pa WHERE pa.prop_id = prop.prop_id  AND pa.prop_atr_type_code = 'PRC'
ORDER BY prop.prop_id, pa.prop_atr_code
SELECT prop_id, MIN(id) as 'start' INTO #mySt2 FROM #myPRCs GROUP BY prop_id
UPDATE #myPRCs set seq = id-M.start FROM #myPRCs r, #mySt2 M WHERE r.prop_id = M.prop_id
DROP TABLE #mySt2
SELECT  nsf_rcvd_date, dd_rcom_date, prop.pgm_annc_id, Dir, prop.org_code, Pgm, PO,
obj_clas_code, prop.natr_rqst_abbr, prop.prop_stts_abbr, ILN, prop.lead, Div, prop.prop_id, pi_last_name, pi_frst_name, inst_name,
prop.prop_titl_txt, prop.rqst_dol, prop.rqst_eff_date, prop.rqst_mnth_cnt, prop.cntx_stmt_id, prop.inst_id, prop.pi_id, st_code,
(SELECT MAX( CASE pa.seq WHEN 0 THEN     rtrim(pa.prop_atr_code) END)+
        MAX( CASE pa.seq WHEN 1 THEN ','+rtrim(pa.prop_atr_code) END)+
        MAX( CASE pa.seq WHEN 2 THEN ','+rtrim(pa.prop_atr_code) END)+ 
        MAX( CASE pa.seq WHEN 3 THEN ','+rtrim(pa.prop_atr_code) END)+
        MAX( CASE pa.seq WHEN 4 THEN ','+rtrim(pa.prop_atr_code) END)+
        MAX( CASE pa.seq WHEN 5 THEN ','+rtrim(pa.prop_atr_code) END)+
        MAX( CASE pa.seq WHEN 6 THEN ','+rtrim(pa.prop_atr_code) END)
        FROM #myPRCs pa WHERE pa.prop_id = prop.prop_id) AS 'PRCs',
CASE WHEN PROP_COVR.HUM_DATE is not NULL THEN convert(varchar(10),PROP_COVR.HUM_DATE,1) WHEN PROP_COVR.humn_date_pend_flag='Y' THEN 'Pend' END AS humn_date,
CASE WHEN PROP_COVR.VERT_DATE is not NULL THEN convert(varchar(10),PROP_COVR.VERT_DATE,1) WHEN PROP_COVR.vrtb_date_pend_flag='Y' THEN 'Pend' END AS vrtb_date,
(SELECT  MAX( CASE mySeq.seq WHEN 0 THEN mySeq.CODE ELSE '' END ) + ' ' +
                 MAX( CASE mySeq.seq WHEN 1 THEN mySeq.CODE ELSE '' END ) + ' ' +
                MAX( CASE mySeq.seq WHEN 2 THEN mySeq.CODE ELSE '' END ) + ' ' +
                MAX( CASE mySeq.seq WHEN 3 THEN mySeq.CODE ELSE '' END )
  FROM (SELECT  r1.TEMP_PROP_ID, r1.CODE,
                            (SELECT  count(*) FROM FLflpdb.flp.routing r2
                              WHERE r2.TEMP_PROP_ID = r1.TEMP_PROP_ID
                                             AND r2.SEQUENCE < r1.SEQUENCE ) as 'seq'
                 FROM FLflpdb.flp.routing r1) mySeq
  WHERE psc.TEMP_PROP_ID = mySeq.TEMP_PROP_ID)  AS 'rout_PECs',
(SELECT MAX(CASE b.seq WHEN 1 THEN b.ctry_name ELSE '' END)+
    MAX(CASE b.seq WHEN 2 THEN '; '+b.ctry_name ELSE '' END)+
    MAX(CASE b.seq WHEN 3 THEN '; '+b.ctry_name ELSE '' END)+
    MAX(CASE b.seq WHEN 4 THEN '; '+b.ctry_name ELSE '' END)+
    MAX(CASE b.seq WHEN 5 THEN '; '+b.ctry_name ELSE '' END)+
    MAX(CASE b.seq WHEN 6 THEN '; '+b.ctry_name ELSE '' END)+
    MAX(CASE b.seq WHEN 7 THEN '; '+b.ctry_name ELSE '' END)+
    MAX(CASE b.seq WHEN 8 THEN '; '+b.ctry_name ELSE '' END) 
    FROM #myCtry b WHERE b.TEMP_PROP_ID = psc.TEMP_PROP_ID) AS Country
FROM #myProps prop
JOIN csd.prop_subm_ctl_vw psc ON prop.prop_id = psc.prop_id 
JOIN FLflpdb.flp.PROP_COVR PROP_COVR ON psc.TEMP_PROP_ID = PROP_COVR.TEMP_PROP_ID
ORDER BY prop.lead, ILN, prop.prop_id

DROP TABLE #myPRCs  DROP TABLE #myProps
--]PD3_props

--[PD3_projSumm
SELECT   nsf_rcvd_date, dd_rcom_date, mp.pgm_annc_id, Dir, mp.org_code,  Pgm,  PO, mp.obj_clas_code, natr_rqst_abbr, prop_stts_abbr,
ILN, mp.Div, mp.prop_id, pi_last_name, pi_frst_name, inst_name, st_code, pi_emai_addr, pi_id, inst_id, 
prop_titl_txt, (SELECT sum(p.rqst_dol) FROM #myProps p where p.lead = mp.lead) AS rqst_dol_tot, mp.rqst_eff_date, mp.rqst_mnth_cnt,
s.SPCL_CHAR_PDF, s.PROJ_SUMM_TXT, s.INTUL_MERT, s.BRODR_IMPT
FROM #myProps mp
JOIN csd.pgm_annc pgm_annc ON mp.pgm_annc_id = pgm_annc.pgm_annc_id
JOIN csd.org org ON mp.org_code = org.org_code
JOIN FLflpdb.flp.obj_clas_pars oc ON mp.obj_clas_code = oc.obj_clas_code
JOIN FLflpdb.flp.prop_subm_ctl psc ON mp.prop_id = psc.PROP_ID
LEFT OUTER JOIN FLflpdb.flp.proj_summ s ON psc.TEMP_PROP_ID = s.TEMP_PROP_ID
WHERE mp.ILN < 'M'
ORDER BY mp.lead
DROP TABLE #myProps
--]PD3_projSumm


--Suggested and unwanted reviewers
--[PD3_SugRev
SELECT prop.nsf_rcvd_date, Dir, org_code, Pgm, PO, prop.natr_rqst_abbr as natr_rqst, prop.prop_stts_abbr as prop_stts,
prop.ILN, prop.lead, Div, prop.prop_id, prop.pi_last_name, prop.pi_frst_name, prop.inst_name,
srevr.revr_want as Suggested_Reviewers, srevr.revr_dont_want as Unwanted_Reviewers, 
prop.prop_titl_txt, len(srevr.revr_want)+len(srevr.revr_dont_want) as len_tot
FROM #myProps prop 
JOIN csd.prop_subm_ctl_vw ctl ON prop.PROP_ID=ctl.PROP_ID 
JOIN csd.sugg_revr_vw srevr ON ctl.TEMP_PROP_ID=srevr.TEMP_PROP_ID 
ORDER BY prop.lead, prop.ILN, prop.prop_id
DROP TABLE #myProps
--]PD3_SugRev


-- all proposals not on panels
--[PD3_Orphans
SELECT nsf_rcvd_date, Dir, org_code, Pgm, PO,  nullif(Count(rev_prop.rev_sent_date),0) as sent_adhoc, nullif(Count(rev_prop.rev_sent_date)-Count(rev_prop.rev_rtrn_date),0) as out_adhoc, MAX(rev_prop.rev_due_date) as last_due,
obj_clas_code, natr_rqst_abbr, prop_stts_abbr, ILN, lead, Div, prop.prop_id, pi_last_name, pi_frst_name, inst_name, prop_titl_txt, prop_stts_txt
FROM #myProps prop
LEFT OUTER JOIN  csd.rev_prop rev_prop ON prop.prop_id = rev_prop.prop_id AND rev_prop.rev_type_code='R'
WHERE not exists (select pp.panl_id FROM csd.panl_prop pp WHERE prop.prop_id = pp.prop_id)
GROUP BY prop_stts_txt, Pgm, PO, natr_rqst_abbr, Dir, ILN, lead, Div, prop.prop_id, pi_last_name, pi_frst_name, inst_name, prop_titl_txt, nsf_rcvd_date, prop_stts_abbr
ORDER BY prop_stts_abbr, lead, ILN
DROP TABLE #myProps
--]PD3_Orphans

-- for the proposals in #myProps, find the change PI transfers and all internal changes of PD.
--[PD3_Transfer
SELECT p.Pgm, p.natr_rqst_abbr, p.prop_stts_abbr, p.ILN, p.lead,p.Div, p.prop_id, a.awd_id
INTO #myPropsAwds
FROM #myProps p  
JOIN csd.amd a ON p.prop_id = a.prop_id 
UNION select p.Pgm, p.natr_rqst_abbr, p.prop_stts_abbr, p.ILN, p.lead, p.Div, p.prop_id, c.PREV_AWD_ID 
from #myProps p 
join csd.prop_subm_ctl_vw ctl on p.prop_id = ctl.prop_id 
join csd.prop_covr_vw c on c.TEMP_PROP_ID = ctl.TEMP_PROP_ID 
WHERE c.PREV_AWD_ID <> '' AND p.natr_rqst_abbr not in ('NEW','RNEW','ABR') 
SELECT distinct a.Pgm, a.natr_rqst_abbr, a.prop_stts_abbr, a.ILN, a.lead, a.Div, a.prop_id, a.awd_id, min(l.last_prop_id) as last_prop_id, pi.pi_last_name, inst.inst_shrt_name, 'awd ' as xfer_type, 
    CASE WHEN l.org_code <> prop.org_code THEN prop.org_code END as fromORG, 
    CASE WHEN l.pgm_ele_code <> prop.pgm_ele_code THEN prop.pgm_ele_code END as fromPEC,
    CASE WHEN l.pm_ibm_logn_id <> prop.pm_ibm_logn_id THEN prop.pm_ibm_logn_id END as fromPA_PO,  
    l.org_code as toORG, l.pgm_ele_code as toPEC, l.pm_ibm_logn_id as toPA_PO, min(l.awd_chg_date) as crtd_date, convert(varchar(255),null) as note
INTO #myXfers
FROM #myPropsAwds a 
JOIN csd.awd_chg_log l ON l.awd_id = a.awd_id
JOIN csd.prop prop ON prop.prop_id = l.last_prop_id
JOIN csd.pi_vw pi ON prop.pi_id = pi.pi_id
JOIN csd.inst inst ON prop.inst_id = inst.inst_id
WHERE a.awd_id <> '' 
Group by a.Pgm, a.prop_stts_abbr, a.natr_rqst_abbr, a.lead, a.ILN, a.Div, a.prop_id, a.awd_id, pi.pi_last_name, inst.inst_shrt_name, l.pm_ibm_logn_id, l.org_code, l.pgm_ele_code, prop.pm_ibm_logn_id, prop.org_code, prop.pgm_ele_code
insert into #myXfers SELECT x.*, 'prop' as xfer_type, 
    CASE WHEN to_org_code <> from_org_code THEN from_org_code END as fromORG, 
    CASE WHEN to_pgm_ele_code <> from_pgm_ele_code THEN from_pgm_ele_code END as fromPEC,
    CASE WHEN to_lan_id <> from_lan_id THEN from_lan_id END as fromPA_PO, 
    to_org_code as toORG, to_pgm_ele_code as toPEC, to_lan_id as toPA_PO, h.crtd_date, isnull(h.note,'') as note 
FROM ( SELECT a.Pgm, a.natr_rqst_abbr, a.prop_stts_abbr,  a.ILN,a.lead, a.Div, a.prop_id, a.awd_id, a.last_prop_id, a.pi_last_name, a.inst_shrt_name
       FROM #myXfers a 
       WHERE a.prop_id <> a.last_prop_id 
 UNION SELECT p.Pgm, p.natr_rqst_abbr, p.prop_stts_abbr,  p.ILN, p.lead, p.Div, p.prop_id, a.awd_id, p.prop_id, p.pi_last_name, p.inst_name 
       FROM #myProps p 
       LEFT OUTER JOIN #myPropsAwds a ON p.prop_id = a.prop_id ) x
JOIN FLflpdb.flp.prop_ownr_hist h ON h.prop_id = x.last_prop_id 
WHERE h.ownr_stts_code = 9 
select  * from #myXfers order by lead, ILN, prop_id, awd_id, crtd_date
DROP TABLE #myProps DROP TABLE #myPropsAwds DROP TABLE #myXfers
--]PD3_Transfer

-- This needs a list of panel ids to pull reviewer information and statistics.
-- todo: add summary statistics
--[PD3_Panl1
SET NOCOUNT ON SELECT pr.revr_id, pr.revr_id, revr.revr_last_name, revr.revr_frst_name, revr.gend_code,
revr.inst_id, inst.inst_shrt_name, inst.st_code, revr.revr_dept_name, f.fos_txt as field, 
COUNT(pr.panl_id) as nPanls, MAX(panl.panl_bgn_date) as latest,
'' AS 'nProps', '' as 'min_score', '' as 'avg_score', '' as 'max_score', '' as 'std_score',
'' AS 'avg_len', '' AS 'std_len', '' AS 'avg_days_early', ''  AS 'std_days_early',
'' as 'nProps_all', '' as 'latest_all', '' as 'min_score_all', '' as 'avg_score_all', '' as 'max_score_all', '' as 'std_score_all',
'' AS 'avg_len_all', '' AS 'std_len_all', '' AS 'avg_days_early_all', ''  AS 'std_days_early_all',
ra.revr_addr_txt as 'revr_email', revr.pi_id, pi.pi_emai_addr
FROM csd.panl_revr pr
JOIN csd.revr revr ON pr.revr_id = revr.revr_id 
JOIN csd.panl panl ON pr.panl_id = panl.panl_id
LEFT OUTER JOIN csd.fos f on f.prmy_fos_code = revr.prmy_fos_code
LEFT OUTER JOIN csd.revr_opt_addr_line ra ON pr.revr_id = ra.revr_id AND ra.addr_lne_type_code='E'
LEFT OUTER JOIN csd.inst inst ON revr.inst_id = inst.inst_id 
LEFT OUTER JOIN csd.pi_vw pi ON revr.pi_id = pi.pi_id
WHERE pr.panl_id IN ('
--]PD3_Panl1
--list panel ids here
--[PD3_Panl2
') GROUP BY pr.revr_id, revr.revr_last_name, revr.revr_frst_name, revr.gend_code, revr.inst_id, inst.inst_shrt_name, inst.st_code, revr.revr_dept_name, f.fos_txt, ra.revr_addr_txt, revr.pi_id, pi.pi_emai_addr
ORDER BY revr.revr_last_name, revr.revr_frst_name
--]PD3_Panl2


-- get reviewers
-- todo: add summary statistics
--[PD3_Rev
SELECT   r.revr_id, COUNT(r.prop_id) as nProps, MAX(r.rev_rtrn_date) as latest, revr.revr_last_name, revr.revr_frst_name, revr.gend_code,
revr.inst_id, inst.inst_shrt_name, inst.st_code, revr.revr_dept_name, f.fos_txt as field, 
'' as 'min_score', '' as 'avg_score', '' as 'max_score', '' as 'std_score',
'' AS 'avg_len', '' AS 'std_len', '' AS 'avg_days_early', ''  AS 'std_days_early',
ra.revr_addr_txt as 'revr_email', revr.pi_id, pi.pi_emai_addr
FROM (SELECT rp.revr_id, rp.prop_id, rp.rev_rtrn_date 
    FROM (SELECT DISTINCT lead FROM #myProps) p
    JOIN csd.rev_prop rp ON rp.prop_id = p.lead
    WHERE rp.rev_rtrn_date is not NULL and rp.rev_type_code <> 'C') r
JOIN csd.revr revr ON revr.revr_id = r.revr_id
LEFT OUTER JOIN csd.fos f on f.prmy_fos_code = revr.prmy_fos_code
LEFT OUTER JOIN csd.revr_opt_addr_line ra ON ra.revr_id = r.revr_id AND ra.addr_lne_type_code='E'
LEFT OUTER JOIN csd.inst inst ON revr.inst_id = inst.inst_id 
LEFT OUTER JOIN csd.pi_vw pi ON revr.pi_id = pi.pi_id
GROUP BY r.revr_id, revr.revr_last_name, revr.revr_frst_name, revr.gend_code, f.fos_txt, revr.inst_id, inst.inst_shrt_name, inst.st_code, revr.revr_dept_name, ra.revr_addr_txt, revr.pi_id, pi.pi_emai_addr
ORDER BY nProps DESC, revr.revr_last_name, revr.revr_frst_name
DROP TABLE #myProps DROP TABLE #revScores
--]PD3_Rev

--results of proposals on panels

--[PD3_PropPanl
SELECT p.lead, rs.string, rs.score, rev_prop_vw.revr_id, rev_prop_vw.rev_rtrn_date, id=identity(18), 0 as 'seq'
INTO #myRevs
FROM #myProps p
JOIN csd.rev_prop_vw rev_prop_vw ON p.lead = rev_prop_vw.prop_id
JOIN csd.rev_prop rev_prop ON p.lead = rev_prop.prop_id AND rev_prop_vw.revr_id = rev_prop.revr_id
JOIN #revScores rs ON rev_prop_vw.rev_prop_rtng_ind = rs.yn 
WHERE p.ILN < 'M' AND rev_prop.rev_stts_code <> 'C'
ORDER BY lead, score DESC
SELECT lead, MIN(id) as 'start' INTO #myStarts FROM #myRevs GROUP BY lead
UPDATE #myRevs set seq = id-M.start FROM #myRevs r, #myStarts M WHERE r.lead = M.lead
DROP TABLE #myStarts
SELECT prop.lead, pp.prop_seq_num as disc_ordr, pp.panl_id, CASE WHEN pps.panl_id IS NOT NULL THEN 'Y' END AS panl_held, 
pps.panl_summ_rlse_flag as summ_rlse, pps.PROP_ORDR as rank, pps.RCOM_SEQ_NUM as rcom_seq, prd.rcom_abbr, prd.rcom_txt
INTO #myPanlProp 
FROM #myProps prop
JOIN csd.panl_prop pp ON prop.lead = pp.prop_id
LEFT OUTER JOIN FLflpdb.flp.panl_prop_summ pps ON prop.lead = pps.PROP_ID AND pp.panl_id = pps.panl_id
LEFT OUTER JOIN FLflpdb.flp.panl_rcom_def prd ON pp.panl_id = prd.panl_id AND pps.RCOM_SEQ_NUM = prd.RCOM_SEQ_NUM
WHERE prop.ILN < 'M' 
SELECT pl.panl_id, panl.panl_name, panl.panl_bgn_date, panl.panl_loc, convert(varchar,SUM(ps.rtCount)) + ' proj: ' +
   MAX( CASE ps.RCOM_SEQ_NUM WHEN 1 THEN        convert(varchar,ps.rtCount) + ' ' +  ps.RCOM_ABBR ELSE '' END ) +
   MAX( CASE ps.RCOM_SEQ_NUM WHEN 2 THEN ', ' + convert(varchar,ps.rtCount) + ' ' +  ps.RCOM_ABBR ELSE '' END ) +
   MAX( CASE ps.RCOM_SEQ_NUM WHEN 3 THEN ', ' + convert(varchar,ps.rtCount) + ' ' +  ps.RCOM_ABBR ELSE '' END ) +
   MAX( CASE ps.RCOM_SEQ_NUM WHEN 4 THEN ', ' + convert(varchar,ps.rtCount) + ' ' +  ps.RCOM_ABBR ELSE '' END ) +
   MAX( CASE ps.RCOM_SEQ_NUM WHEN 5 THEN ', ' + convert(varchar,ps.rtCount) + ' ' +  ps.RCOM_ABBR ELSE '' END ) +
   MAX( CASE ps.RCOM_SEQ_NUM WHEN 6 THEN ', ' + convert(varchar,ps.rtCount) + ' ' +  ps.RCOM_ABBR ELSE '' END ) as 'panlSumm'
INTO #myPanls
FROM (SELECT DISTINCT panl_id FROM #myPanlProp) pl
JOIN csd.panl panl ON pl.panl_id = panl.panl_id
LEFT OUTER JOIN (SELECT p2.panl_id, panl_rcom_def.RCOM_SEQ_NUM, panl_rcom_def.RCOM_ABBR, Count(panl_prop_summ.PROP_ID) AS rtCount
    FROM (SELECT DISTINCT panl_id FROM #myPanlProp) p2
    JOIN FLflpdb.flp.panl_prop_summ panl_prop_summ ON p2.panl_id = panl_prop_summ.PANL_ID
    JOIN FLflpdb.flp.panl_rcom_def panl_rcom_def ON panl_prop_summ.PANL_ID = panl_rcom_def.PANL_ID AND panl_prop_summ.RCOM_SEQ_NUM = panl_rcom_def.RCOM_SEQ_NUM
    JOIN csd.prop pr ON panl_prop_summ.PROP_ID = pr.prop_id AND pr.prop_id=isnull(pr.lead_prop_id,pr.prop_id) 
    GROUP BY p2.panl_id, panl_rcom_def.RCOM_SEQ_NUM, panl_rcom_def.RCOM_ABBR ) ps ON pl.panl_id = ps.panl_id
GROUP BY pl.panl_id, panl.panl_name, panl.panl_bgn_date, panl.panl_loc 
SELECT  prop.*, pp.panl_id, pn.panl_bgn_date, pn.panl_loc, pp.disc_ordr, pn.panl_name, 
pp.panl_held, pp.summ_rlse, pp.rank, pp.rcom_seq, pp.rcom_abbr, pp.rcom_txt,
CASE WHEN pp.rank IS NOT NULL THEN 'rk '+convert(varchar,pp.rank)+' of ' ELSE '' END + pn.panlSumm AS rankInPanl, 
CASE WHEN PROP_COVR.HUM_DATE is not NULL THEN convert(varchar(10),PROP_COVR.HUM_DATE,1) WHEN PROP_COVR.humn_date_pend_flag='Y' THEN 'Pend' END AS humn_date,
CASE WHEN PROP_COVR.VERT_DATE is not NULL THEN convert(varchar(10),PROP_COVR.VERT_DATE,1) WHEN PROP_COVR.vrtb_date_pend_flag='Y' THEN 'Pend' END AS vrtb_date,
revs.reviews, revs.avg_score, revs.last_rev_date,
(SELECT MAX(CASE b.seq WHEN 1 THEN b.ctry_name ELSE '' END)+
    MAX(CASE b.seq WHEN 2 THEN '; '+b.ctry_name ELSE '' END)+
    MAX(CASE b.seq WHEN 3 THEN '; '+b.ctry_name ELSE '' END)+
    MAX(CASE b.seq WHEN 4 THEN '; '+b.ctry_name ELSE '' END)+
    MAX(CASE b.seq WHEN 5 THEN '; '+b.ctry_name ELSE '' END)+
    MAX(CASE b.seq WHEN 6 THEN '; '+b.ctry_name ELSE '' END)+
    MAX(CASE b.seq WHEN 7 THEN '; '+b.ctry_name ELSE '' END)+
    MAX(CASE b.seq WHEN 8 THEN '; '+b.ctry_name ELSE '' END) 
    FROM (SELECT prop_spcl_item_vw.TEMP_PROP_ID, prop_spcl_item_vw.SPCL_ITEM_CODE, ctry.ctry_name, 
            (SELECT COUNT(*) FROM csd.prop_spcl_item_vw sp2
             JOIN csd.ctry ctry2 ON sp2.SPCL_ITEM_CODE = ctry2.ctry_code AND ctry.ctry_code >= ctry2.ctry_code
             WHERE sp2.TEMP_PROP_ID = prop_spcl_item_vw.TEMP_PROP_ID) as seq
          FROM csd.prop_spcl_item_vw prop_spcl_item_vw
          JOIN csd.ctry ctry ON prop_spcl_item_vw.SPCL_ITEM_CODE = ctry.ctry_code ) b 
     WHERE b.TEMP_PROP_ID = prop_subm_ctl_vw.TEMP_PROP_ID
     GROUP BY b.TEMP_PROP_ID) AS country 
FROM #myProps prop
JOIN csd.prop_subm_ctl_vw prop_subm_ctl_vw ON prop.prop_id = prop_subm_ctl_vw.PROP_ID
JOIN FLflpdb.flp.PROP_COVR PROP_COVR ON prop_subm_ctl_vw.TEMP_PROP_ID = PROP_COVR.TEMP_PROP_ID
LEFT OUTER JOIN #myPanlProp pp ON prop.lead = pp.lead
LEFT OUTER JOIN #myPanls pn ON pp.panl_id = pn.panl_id 
LEFT OUTER JOIN (SELECT r.lead, AVG(r.score) AS avg_score, MAX(r.rev_rtrn_date) AS last_rev_date, MAX(CASE r.seq WHEN 0 THEN r.string ELSE '' END)+
      MAX(CASE r.seq WHEN 1 THEN ','+r.string ELSE '' END)+
      MAX(CASE r.seq WHEN 2 THEN ','+r.string ELSE '' END)+
      MAX(CASE r.seq WHEN 3 THEN ','+r.string ELSE '' END)+
      MAX(CASE r.seq WHEN 4 THEN ','+r.string ELSE '' END)+
      MAX(CASE r.seq WHEN 5 THEN ','+r.string ELSE '' END)+
      MAX(CASE r.seq WHEN 6 THEN ','+r.string ELSE '' END)+
      MAX(CASE r.seq WHEN 7 THEN ','+r.string ELSE '' END)+
      MAX(CASE r.seq WHEN 8 THEN ','+r.string ELSE '' END)+
      MAX(CASE r.seq WHEN 9 THEN ','+r.string ELSE '' END) AS reviews
    FROM #myRevs r 
    GROUP BY r.lead) revs ON prop.lead = revs.lead
ORDER BY pn.panl_bgn_date, pn.panl_id, pp.disc_ordr, prop.lead, prop.ILN
DROP TABLE #myProps DROP TABLE #myPanls DROP TABLE #revScores DROP TABLE #myRevs
--]PD3_PropPanl


--[PD3_NewInst
SELECT p.nsf_rcvd_date,  p.Dir, p.org_code, Pgm, PO, newInst.inst_id, inst.inst_name, inst.st_code, inst.ctry_code,  
inst.awd_perf_inst_code, inst.duns_id, inst.perf_org_code, perf_org.perf_org_txt as perf_org_type, 
prop_stts_abbr, natr_rqst_abbr, Div, p.prop_id, pi_last_name, pi_frst_name, inst.inst_shrt_name,  
prop_stts_txt, p.prop_titl_txt
FROM (SELECT pendInst.inst_id
    FROM (SELECT distinct prop.inst_id FROM #myProps prop WHERE prop.prop_stts_abbr in ('PEND','RCOM') ) pendInst
    WHERE NOT EXISTS (SELECT * FROM csd.awd awd WHERE awd.AWD_EXP_DATE >= dateadd(mm, -58, getdate()) AND awd.INST_ID = pendInst.INST_ID ) ) newInst
LEFT OUTER JOIN csd.inst inst ON inst.inst_id = newInst.inst_id
LEFT OUTER JOIN csd.perf_org perf_org ON perf_org.perf_org_code = inst.perf_org_code 
JOIN #myProps p ON p.inst_id = newInst.inst_id 
ORDER BY inst_name
DROP TABLE #myProps
--]PD3_NewInst


--[PD3_BulkCOIrevs
SELECT DISTINCT revr.revr_last_name, revr.revr_frst_name, a.revr_addr_txt AS 'email'
FROM csd.rev_prop rp
JOIN csd.revr revr ON rp.revr_id = revr.revr_id
LEFT OUTER JOIN csd.revr_opt_addr_line a ON rp.revr_id = a.revr_id AND a.addr_lne_type_code='E'
 WHERE rp.prop_id IN ('1612999','1613905','1614150','1614385','1614562','1614584','1615458','1615584','1615704','1615845','1616248','1617166','1617193','1617256','1617354','1617461','1617546','1617605','1617617','1617641','1617735','1617819','1617830','1617951','1618034','1618116','1618345','1618380','1618469','1618501','1618593','1618657','1618717','1618818','1618866','1618894','1618900','1618981','1619064','1619095','1619107','1619144','1619256','1619257','1619325','1619330','1619345','1619350','1619402','1623212','1624382','1629809','1636571','1637380','1637534','1637566','1639912','1640970','1644747','1649473','1651569','1651932','1651987','1651999','1652218','1652640','1652695','1652857','1652909','1653504','1654106','1654261','1654850','1655367','1655422','1656905','1657069','1657147','1657316','1657325','1657472','1657939','1659001','1660636','1665252')
 ORDER BY revr.revr_last_name, revr.revr_frst_name
--]PD3_BulkCOIrevs

--[PD3_BulkCOIprops
SELECT Div, prop_id, ILN, pi_last_name, inst_name FROM #myProps prop WHERE ILN < 'M' AND prop_stts_abbr NOT IN ('WTH','RTNR') ORDER BY lead, ILN, prop_id
DROP TABLE #myProps
--]PD3_BulkCOIprops

--For proposals with non-zero subaward amounts, see if the subawards add up to the right value
--[PD3_SubAwd
SELECT p.prop_id, psc.temp_prop_id 
INTO #hasSubAwd
FROM #myProps p 
JOIN csd.prop_subm_ctl_vw psc ON p.prop_id = psc.prop_id 
WHERE EXISTS (SELECT * FROM csd.budg_vw b WHERE psc.TEMP_PROP_ID = b.TEMP_PROP_ID AND b.sub_ctr_req_dol>0)  
SELECT prop_id, perf_inst_id, revn_num, round(sum(subAwd),0) AS subAwd, round(sum(bud_tot),0) AS bud_tot  
INTO #myBudg
FROM ( 
  SELECT p.prop_id, b.perf_inst_id, b.revn_num, isnull(b.sub_ctr_req_dol,0) as subAwd, 
    isnull(OTHR_SR_REQ_DOL,0)+isnull(PDOC_REQ_DOL,0)+isnull(OTH_PROF_REQ_DOL,0)+isnull(GRAD_REQ_DOL,0)+isnull(UN_GRAD_REQ_DOL,0)+ 
    isnull(SEC_REQ_DOL,0)+isnull(OTH_PERS_REQ_DOL,0)+isnull(FRIN_BNFT_REQ_DOL,0)+isnull(DOM_TRAV_REQ_DOL,0)+isnull(FRGN_TRAV_REQ_DOL,0)+ 
    isnull(PART_SUPT_STPD_DOL,0)+isnull(PART_SUPT_TRAV_DOL,0)+isnull(PART_SUPT_SUBS_DOL,0)+isnull(PART_SUPT_OTH_DOL,0)+isnull(MATL_REQ_DOL,0)+ 
    isnull(PUB_REQ_DOL,0)+isnull(CNSL_REQ_DOL,0)+isnull(CPTR_SERV_REQ_DOL,0)+isnull(SUB_CTR_REQ_DOL,0)+isnull(OTH_DRCT_CST_REQ_DOL,0)+isnull(RSID_REQ_DOL,0) AS bud_tot 
   FROM #hasSubAwd p 
   JOIN csd.budg_vw b ON p.TEMP_PROP_ID = b.TEMP_PROP_ID 
  UNION ALL SELECT p.prop_id, sr.perf_inst_id, sr.revn_num, 0, sr_req_dol 
   FROM #hasSubAwd p
   JOIN csd.sr_pers_resc_vw sr ON p.TEMP_PROP_ID = sr.TEMP_PROP_ID 
  UNION ALL SELECT p.prop_id, e.perf_inst_id, e.revn_num, 0, isnull(e.equp_cst_dol_req,0)
   FROM #hasSubAwd p 
   JOIN csd.equp_cst_vw e ON p.TEMP_PROP_ID = e.TEMP_PROP_ID 
  UNION ALL SELECT p.prop_id, i.perf_inst_id, i.revn_num, 0, round(isnull(i.IDIR_CST_RATE*i.idir_cst_dol_req/100,0),0)
   FROM #hasSubAwd p
   JOIN csd.idir_cst_vw i ON p.TEMP_PROP_ID = i.TEMP_PROP_ID) b
GROUP BY prop_id, perf_inst_id, revn_num  
ORDER BY prop_id, subAwd DESC, bud_tot DESC 
CREATE INDEX myBudg_ix ON #myBudg(prop_id) 
SELECT  nsf_rcvd_date, dd_rcom_date, Dir, prop.org_code, Pgm, PO, prop.natr_rqst_abbr AS natr_rqst, prop.prop_stts_abbr AS prop_stts,
prop.ILN, prop.lead, Div, prop.prop_id, prop.pi_last_name, prop.inst_name,
 x.inst_shrt_name AS perf_inst, x.revn_num, x.bud_tot AS budg_total, x.budg_diff, x.subAwd, sub_diff as subAwd_diff,  
prop.prop_titl_txt
FROM #myProps prop
JOIN (SELECT prop_id, inst.inst_shrt_name, revn_num,  
    bud_tot,  NULLIF(bud_tot - (SELECT SUM(b.budg_tot_dol) FROM csd.eps_blip b WHERE b.prop_id = p.prop_id AND b.revn_num = 0),0) AS budg_diff,  
    subAwd, NULLIF(subAwd - (SELECT ISNULL(SUM(bud_tot),0) FROM #myBudg s WHERE subAwd = 0 AND s.prop_id = p.prop_id AND s.revn_num = p.revn_num),0) AS sub_diff  
    FROM #myBudg p 
JOIN csd.inst inst ON p.perf_inst_id = inst.inst_id 
WHERE subAwd > 0) x  ON prop.prop_id = x.prop_id
ORDER BY prop.lead, prop.ILN, prop.prop_id
DROP TABLE #myProps DROP TABLE #hasSubAwd DROP TABLE #myBudg
--]PD3_SubAwd



--[PD3_panls
SELECT panl_prop.panl_id, panl.panl_name, panl.pm_logn_id, Count(prop.lead) AS numProps, panl.panl_bgn_date, panl.panl_loc, panl.pgm_ele_code, panl.fund_org_code, panl.fund_pgm_ele_code
INTO #myPanls FROM #myProps prop
JOIN csd.panl_prop panl_prop ON prop.prop_id = panl_prop.prop_id 
JOIN csd.panl panl ON panl_prop.panl_id = panl.panl_id
WHERE prop.ILN < 'M' 
GROUP BY panl_prop.panl_id, panl.panl_name, panl.pm_logn_id, panl.panl_bgn_date, panl.panl_loc, panl.pgm_ele_code, panl.fund_org_code, panl.fund_pgm_ele_code
SELECT  mp.*, convert(varchar,SUM(ps.rtCount)) + ' rated projects: ' +
   MAX( CASE ps.RCOM_SEQ_NUM WHEN 1 THEN        convert(varchar,ps.rtCount) + ' ' +  ps.RCOM_ABBR ELSE '' END ) +
   MAX( CASE ps.RCOM_SEQ_NUM WHEN 2 THEN ', ' + convert(varchar,ps.rtCount) + ' ' +  ps.RCOM_ABBR ELSE '' END ) +
   MAX( CASE ps.RCOM_SEQ_NUM WHEN 3 THEN ', ' + convert(varchar,ps.rtCount) + ' ' +  ps.RCOM_ABBR ELSE '' END ) +
   MAX( CASE ps.RCOM_SEQ_NUM WHEN 4 THEN ', ' + convert(varchar,ps.rtCount) + ' ' +  ps.RCOM_ABBR ELSE '' END ) +
   MAX( CASE ps.RCOM_SEQ_NUM WHEN 5 THEN ', ' + convert(varchar,ps.rtCount) + ' ' +  ps.RCOM_ABBR ELSE '' END ) +
   MAX( CASE ps.RCOM_SEQ_NUM WHEN 6 THEN ', ' + convert(varchar,ps.rtCount) + ' ' +  ps.RCOM_ABBR ELSE '' END ) as 'panlSumm'
FROM #myPanls mp
LEFT OUTER JOIN (SELECT pl.panl_id,  panl_rcom_def.RCOM_SEQ_NUM, panl_rcom_def.RCOM_ABBR, Count(panl_prop_summ.PROP_ID) AS rtCount
    FROM (SELECT DISTINCT panl_id FROM #myPanls) pl
    JOIN FLflpdb.flp.panl_prop_summ panl_prop_summ ON pl.panl_id = panl_prop_summ.PANL_ID
    JOIN FLflpdb.flp.panl_rcom_def panl_rcom_def ON panl_prop_summ.PANL_ID = panl_rcom_def.PANL_ID AND panl_prop_summ.RCOM_SEQ_NUM = panl_rcom_def.RCOM_SEQ_NUM
    JOIN csd.prop pr ON panl_prop_summ.PROP_ID = pr.prop_id AND pr.prop_id=isnull(pr.lead_prop_id,pr.prop_id) 
    GROUP BY pl.panl_id, panl_rcom_def.RCOM_SEQ_NUM, panl_rcom_def.RCOM_ABBR ) ps ON mp.panl_id = ps.panl_id
GROUP BY mp.panl_id  ORDER BY mp.panl_bgn_date
DROP TABLE #myProps DROP TABLE #myPanls
--]PD3_panls

