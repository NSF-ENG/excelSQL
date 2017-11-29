set nocount on

SELECT isnull(prop.lead_prop_id,prop.prop_id) AS 'lead_id',
CASE WHEN prop.lead_prop_id IS NULL THEN 'I' WHEN prop.lead_prop_id = prop.prop_id THEN 'L' ELSE 'N' END AS ILN, 
prop.prop_id,nullif(prop.dd_rcom_date,'1/1/1900') AS dd_rcom_date
WHERE (((1=1) AND prop.nsf_rcvd_date >=  '2016-10-01'
 AND (prop.pgm_annc_id = 'NSF 17-537' ) AND (prop.pm_ibm_logn_id = 'jsnoeyi' ) 
AND NOT EXISTS (SELECT * FROM csd.prop_atr pa WHERE pa.prop_id=prop.prop_id AND pa.prop_atr_code = '9150' AND pa.prop_atr_type_code='PRC')
) ) 
ORDER BY lead_id,ILN
CREATE INDEX myProp_ix ON #myProp(prop_id)

--gppr datediff(nullif(pgm_annc.,nsf_rcvd_date),nullif(dd_rcom_date,getdate())

SELECT isnull(prop.lead_prop_id,prop.prop_id) AS 'lead_id',
CASE WHEN prop.lead_prop_id IS NULL THEN 'I' WHEN prop.lead_prop_id = prop.prop_id THEN 'L' ELSE 'N' END AS ILN, prop.prop_id,
(SELECT MAX( CASE pa.prop_atr_seq WHEN 1 THEN pa.prop_atr_code ELSE '' END ) + ' ' +
        MAX( CASE pa.prop_atr_seq WHEN 2 THEN pa.prop_atr_code ELSE '' END ) + ' ' +
        MAX( CASE pa.prop_atr_seq WHEN 3 THEN pa.prop_atr_code ELSE '' END ) + ' ' +
        MAX( CASE pa.prop_atr_seq WHEN 4 THEN pa.prop_atr_code ELSE '' END ) + ' ' +
        MAX( CASE pa.prop_atr_seq WHEN 5 THEN pa.prop_atr_code ELSE '' END ) + ' ' +
        MAX( CASE pa.prop_atr_seq WHEN 6 THEN pa.prop_atr_code ELSE '' END )
        FROM csd.prop_atr pa  WHERE pa.prop_id = prop.prop_id  AND pa.prop_atr_type_code = 'PRC' ) AS 'PRCs',
(SELECT Count(*) FROM csd.rev_prop rp WHERE prop.prop_id=rp.prop_id AND rp.rev_type_code='R')  AS 'adhoc_reqd',
(SELECT Count(*) FROM csd.rev_prop rp WHERE prop.prop_id=rp.prop_id AND rp.rev_type_code='R' AND rp.rev_stts_code='P')  AS 'adhoc_pend',
(SELECT Max(rp.rev_due_date) FROM csd.rev_prop rp WHERE prop.prop_id=rp.prop_id AND rp.rev_type_code='R' AND rp.rev_stts_code='P')  AS 'last_adhoc_due',
(SELECT Count(*) FROM csd.rev_prop rp WHERE prop.prop_id=rp.prop_id AND rp.rev_stts_code='R')  AS 'revRcvd',
(SELECT Max(rp.rev_due_date) FROM csd.rev_prop rp WHERE prop.prop_id=rp.prop_id)  AS 'last_rev_due', 
prop.nsf_rcvd_date,
nullif(prop.dd_rcom_date,'1/1/1900') AS dd_rcom_date
INTO #myProp
FROM csd.prop prop
JOIN csd.prop_stts ps on ps.prop_stts_code=prop.prop_stts_code
JOIN csd.natr_rqst nr on nr.natr_rqst_code = prop.natr_rqst_code
JOIN csd.org  as og on og.org_code=prop.org_code


SELECT panl_prop.prop_id, panl_prop.panl_id, panl.panl_bgn_date, a.rcom_seq_num, b.rcom_abbr, a.prop_ordr
INTO #myPanl
FROM #myProp prop, 
JOIN csd.panl_prop panl_prop ON prop.prop_id=panl_prop.prop_id 
JOIN csd.panl panl ON panl_prop.panl_id = panl.panl_id 
LEFT JOIN flflpdb.flp.panl_prop_summ a ON panl_prop.panl_id = a.panl_id 
LEFT JOIN flflpdb.flp.panl_rcom_def b ON a.panl_id = b.panl_id AND a.rcom_seq_num = b.rcom_seq_num
WHERE  AND 
AND  AND prop.prop_id *= a.prop_id 
CREATE INDEX myPanl_ix ON #myPanl(prop_id)

SELECT getdate() as run_date,mp.*, prop.pgm_annc_id, prop.org_code, prop.pgm_ele_code, prop.pm_ibm_logn_id, 
prop_stts.prop_stts_abbr, prop.prop_stts_code, prop_stts.prop_stts_txt, 
pi.pi_last_name, pi.pi_frst_name, pi.pi_gend_code, inst.inst_shrt_name AS inst_name, 
inst.st_code, prop.prop_titl_txt, natr_rqst.natr_rqst_txt, natr_rqst.natr_rqst_abbr, 
prop.bas_rsch_pct, prop.cntx_stmt_id,
first.panl_id as 'first_panl', first.panl_bgn_date as 'fp_begin', first.rcom_seq_num as 'fp_recno', first.rcom_abbr as 'fp_rec', first.prop_ordr as 'fp_rank', 
last.panl_id as 'last_panl', last.panl_bgn_date as 'lp_begin', last.rcom_seq_num as 'lp_recno', last.rcom_abbr as 'lp_rec', last.prop_ordr as 'lp_rank',
bs.split_tot_dol, bs.split_frwd_date, bs.split_aprv_date,
prop.rqst_dol, prop.rqst_mnth_cnt, nullif(prop.rcom_mnth_cnt,0) AS 'rcom_mnth_cnt', prop.rqst_eff_date, 
nullif(prop.rcom_eff_date,'1900-01-01') AS 'rcom_eff_date', nullif(prop.pm_asgn_date,'1900-01-01') AS pm_asgn_date, nullif(prop.pm_rcom_date,'1900-01-01') AS  pm_rcom_date, 
nullif(prop.dd_rcom_date,'1900-01-01') AS  dd_rcom_date,
awd.awd_id, awd.tot_intn_awd_amt, pi2.pi_last_name, pi2.pi_frst_name, inst2.inst_shrt_name AS inst_awd, awd.awd_titl_txt, awd.pm_ibm_logn_id, awd.org_code, awd.pgm_ele_code, 
awd.pgm_div_code, awd.awd_istr_code, awd.awd_stts_code, awd.fpr_stts_code, awd.awd_stts_date, awd.awd_eff_date, 
awd.awd_exp_date, awd.awd_fin_clos_date, awd.fpr_stts_updt_date, awd.est_fnl_exp_date
FROM  #myProp mp, 
csd.prop prop, 
csd.inst inst, 
csd.natr_rqst 
natr_rqst, 
csd.pi pi, 
csd.prop_stts prop_stts,  
csd.awd awd, 
csd.inst inst2, 
csd.pi pi2,
(SELECT *  FROM #myPanl pn
WHERE  pn.panl_bgn_date =(SELECT min(p.panl_bgn_date) FROM #myPanl p  WHERE pn.prop_id=p.prop_id ) ) first,
(SELECT *  FROM #myPanl pn
WHERE  pn.panl_bgn_date >(SELECT min(p.panl_bgn_date) FROM #myPanl p WHERE pn.prop_id=p.prop_id )
AND pn.panl_bgn_date = (SELECT max(p.panl_bgn_date) FROM #myPanl p  WHERE pn.prop_id=p.prop_id ) ) last,
(SELECT budg_splt.prop_id, Sum(budg_splt.budg_splt_tot_dol) AS 'split_tot_dol', Max(budg_splt.frwd_date) AS 'split_frwd_date', Max(budg_splt.aprv_date) AS 'split_aprv_date'
FROM csd.budg_splt budg_splt GROUP BY budg_splt.prop_id) bs
WHERE mp.prop_id = prop.prop_id AND prop.natr_rqst_code = natr_rqst.natr_rqst_code AND prop.prop_stts_code = prop_stts.prop_stts_code AND prop.inst_id = inst.inst_id AND prop.pi_id = pi.pi_id AND
mp.prop_id *= first.prop_id AND mp.prop_id *= last.prop_id AND mp.prop_id *= bs.prop_id AND mp.prop_id *= awd.awd_id AND awd.inst_id *= inst2.inst_id AND awd.pi_id *= pi2.pi_id
ORDER BY mp.lead_id, mp.ILN, mp.prop_id
drop table #myProp drop table #myPanl