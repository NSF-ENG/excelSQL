-- proposal, split, award tracker  by Maria Loera and Jack Snoeyink, Nov 2017

-- Proposal tracker   
-- Note: this is not as careful about things like ignoring conflicted reviews received
-- todo: sequence PRCs
--[propTrack1
set nocount on
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
(SELECT Max(rp.rev_due_date) FROM csd.rev_prop rp WHERE prop.prop_id=rp.prop_id)  AS 'last_rev_due', prop.nsf_rcvd_date,nullif(prop.dd_rcom_date,'1/1/1900') AS dd_rcom_date
INTO #myProp
FROM csd.prop prop
JOIN csd.prop_stts ps on ps.prop_stts_code=prop.prop_stts_code
JOIN csd.natr_rqst nr on nr.natr_rqst_code = prop.natr_rqst_code
JOIN csd.org  as org on org.org_code=prop.org_code
WHERE ((1=1)
--]propTrack1 AND prop.nsf_rcvd_date >= {ts '2016-10-01 00:00:00'}
 AND ps.prop_stts_abbr = 'AWD'
 AND prop.prop_titl_txt LIKE '%BSF%'
) ) 
--[propTrack2
ORDER BY lead_id,ILN
CREATE INDEX myProp_ix ON #myProp(prop_id)

SELECT panl_prop.prop_id, panl_prop.panl_id, panl.panl_bgn_date, a.rcom_seq_num, b.rcom_abbr, a.prop_ordr
INTO #myPanl
FROM #myProp prop, csd.panl_prop panl_prop, csd.panl panl, flflpdb.flp.panl_prop_summ a, flflpdb.flp.panl_rcom_def b
WHERE prop.prop_id=panl_prop.prop_id AND panl_prop.panl_id = panl.panl_id
AND  panl_prop.panl_id *= a.panl_id AND prop.prop_id *= a.prop_id AND a.panl_id *= b.panl_id  AND  a.rcom_seq_num *= b.rcom_seq_num
CREATE INDEX myPanl_ix ON #myPanl(prop_id)

SELECT getdate() as run_date,mp.*, prop.pgm_annc_id, prop.org_code, prop.pgm_ele_code, prop.pm_ibm_logn_id, prop_stts.prop_stts_abbr, prop.prop_stts_code, prop_stts.prop_stts_txt, pi.pi_last_name, pi.pi_frst_name, pi.pi_gend_code, inst.inst_shrt_name AS inst_name, inst.st_code, prop.prop_titl_txt, natr_rqst.natr_rqst_txt, natr_rqst.natr_rqst_abbr, prop.bas_rsch_pct, prop.cntx_stmt_id,
first.panl_id as 'first_panl', first.panl_bgn_date as 'fp_begin', first.rcom_seq_num as 'fp_recno', first.rcom_abbr as 'fp_rec', first.prop_ordr as 'fp_rank', last.panl_id as 'last_panl', last.panl_bgn_date as 'lp_begin', last.rcom_seq_num as 'lp_recno', last.rcom_abbr as 'lp_rec', last.prop_ordr as 'lp_rank',
bs.split_tot_dol, bs.split_frwd_date, bs.split_aprv_date,
prop.rqst_dol, prop.rqst_mnth_cnt, nullif(prop.rcom_mnth_cnt,0) AS 'rcom_mnth_cnt', prop.rqst_eff_date, nullif(prop.rcom_eff_date,'1900-01-01') AS 'rcom_eff_date', nullif(prop.pm_asgn_date,'1900-01-01') AS pm_asgn_date, nullif(prop.pm_rcom_date,'1900-01-01') AS  pm_rcom_date, nullif(prop.dd_rcom_date,'1900-01-01') AS  dd_rcom_date,
awd.awd_id, awd.tot_intn_awd_amt, pi2.pi_last_name, pi2.pi_frst_name, inst2.inst_shrt_name AS inst_awd, awd.awd_titl_txt, awd.pm_ibm_logn_id, awd.org_code, awd.pgm_ele_code, awd.pgm_div_code, awd.awd_istr_code, awd.awd_stts_code, awd.fpr_stts_code, awd.awd_stts_date, awd.awd_eff_date, awd.awd_exp_date, awd.awd_fin_clos_date, awd.fpr_stts_updt_date, awd.est_fnl_exp_date
FROM  #myProp mp, csd.prop prop, csd.inst inst, csd.natr_rqst natr_rqst, csd.pi pi, csd.prop_stts prop_stts,  csd.awd awd, csd.inst inst2, csd.pi pi2,
(SELECT *  FROM #myPanl pn
WHERE  pn.panl_bgn_date =(SELECT min(p.panl_bgn_date) FROM #myPanl p  WHERE pn.prop_id=p.prop_id ) ) first,
(SELECT *  FROM #myPanl pn
WHERE pn.panl_bgn_date >(SELECT min(p.panl_bgn_date) FROM #myPanl p WHERE pn.prop_id=p.prop_id )
AND pn.panl_bgn_date = (SELECT max(p.panl_bgn_date) FROM #myPanl p  WHERE pn.prop_id=p.prop_id ) ) last,
(SELECT budg_splt.prop_id, Sum(budg_splt.budg_splt_tot_dol) AS 'split_tot_dol', Max(budg_splt.frwd_date) AS 'split_frwd_date', Max(budg_splt.aprv_date) AS 'split_aprv_date'
FROM csd.budg_splt budg_splt GROUP BY budg_splt.prop_id) bs
WHERE mp.prop_id = prop.prop_id AND prop.natr_rqst_code = natr_rqst.natr_rqst_code AND prop.prop_stts_code = prop_stts.prop_stts_code AND prop.inst_id = inst.inst_id AND prop.pi_id = pi.pi_id AND
mp.prop_id *= first.prop_id AND mp.prop_id *= last.prop_id AND mp.prop_id *= bs.prop_id AND mp.prop_id *= awd.awd_id AND awd.inst_id *= inst2.inst_id AND awd.pi_id *= pi2.pi_id
ORDER BY mp.lead_id, mp.ILN, mp.prop_id
drop table #myPanl
--]propTrack2
drop table #myProp

-- Split Tracker

SET NOCOUNT ON 
CREATE TABLE #AddBudgSplts(
prop_id char(7),
budg_yr smallint null,
splt_id char(2) null
) 
CREATE TABLE #OmitBudgSplts(
prop_id char(7),
budg_yr smallint null,
splt_id char(2) null
) 
SELECT b.prop_id,b.budg_yr,b.splt_id,b.budg_splt_tot_dol, b.org_code as Bdg_Org_Code,
b.pgm_ele_code+' - '+pe.pgm_ele_name as PEC_bdg_splt
INTO #myBSplit
from csd.budg_splt b
JOIN csd.prop p on p.prop_id=b.prop_id
JOIN csd.pgm_ele  as pe ON pe.pgm_ele_code=b.pgm_ele_code
JOIN csd.prop_stts ps on ps.prop_stts_code=p.prop_stts_code
JOIN csd.natr_rqst nr on nr.natr_rqst_code = p.natr_rqst_code
JOIN csd.awd_istr ai on p.rcom_awd_istr = ai.awd_istr_code
JOIN csd.org  as og on og.org_code=p.org_code
WHERE (((1=1)
AND p.nsf_rcvd_date >= '2012-10-01'
 AND b.pgm_ele_code ='7797'
 AND ps.prop_stts_abbr NOT IN ('DECL','WTH')
))

INSERT INTO #myBSplit 
select  bs.prop_id,
bs.budg_yr, bs.splt_id,
bs.budg_splt_tot_dol, bs.org_code as Bdg_Org_Code,bs.pgm_ele_code+' - '+pe.pgm_ele_name as PEC_bdg_splt
from #AddBudgSplts t
JOIN CSD.budg_splt bs ON bs.prop_id=t.prop_id  
and isnull(t.budg_yr,bs.budg_yr) = bs.budg_yr 
and isnull(t.splt_id,bs.splt_id) = bs.splt_id 
JOIN csd.pgm_ele  as pe ON pe.pgm_ele_code=bs.pgm_ele_code

delete  from #myBSplit
from  #myBSplit b
Join  #OmitBudgSplts as t
ON b.prop_id=t.prop_id 
and isnull(t.budg_yr,b.budg_yr) = b.budg_yr 
and isnull(t.splt_id,b.splt_id) = b.splt_id 

SELECT bpr.prop_id,bpr.budg_yr,bpr.splt_id, bpr.pgm_ref_code,id=identity(18), 0 as 'seq'
INTO #myBudgPRCs FROM #myBSplit mbs,csd.budg_pgm_ref  bpr
WHERE mbs.prop_id = bpr.prop_id  AND mbs.splt_id= bpr.splt_id AND mbs.budg_yr=bpr.budg_yr
order by bpr.prop_id,bpr.budg_yr,bpr.splt_id, bpr.pgm_ref_code
SELECT prop_id,budg_yr,splt_id, MIN(id) as 'start'
INTO #mySt2 
FROM #myBudgPRCS 
GROUP BY prop_id,budg_yr,splt_id
UPDATE #myBudgPRCs set seq = id-M.start FROM #myBudgPRCs rb, #mySt2 M
WHERE rb.prop_id = M.prop_id AND rb.budg_yr=M.budg_yr AND rb.splt_id= M.splt_id 

SELECT p.prop_id, pa.prop_atr_code,id=identity(18), 0 as 'seq'
INTO #myPropPRCs
FROM (select distinct prop_id from #myBSplit mbs) as p, csd.prop_atr pa
WHERE pa.prop_id = p.prop_id  AND pa.prop_atr_type_code = 'PRC'
order by p.prop_id, pa.prop_atr_code
SELECT prop_id, MIN(id) as 'start'
INTO #mySt3 
FROM #myPropPRCs group by prop_id
UPDATE #myPropPRCs set seq = id-M.start FROM #myPropPRCs r, #mySt3 M
WHERE r.prop_id = M.prop_id 

select p.prop_id,isnull(p.lead_prop_id,p.prop_id) AS lead_id,
CASE WHEN p.lead_prop_id IS NULL THEN 'I' WHEN p.lead_prop_id = p.prop_id THEN 'L' ELSE 'N' END AS ILN,
p.nsf_rcvd_date, p.dd_rcom_date, p.pgm_annc_id,p.org_code as Prop_Org_Code,
p.pgm_ele_code+' - '+pe.pgm_ele_name as PEC_prop,
p.pm_ibm_logn_id as Prop_Pm_ibm_logn_id,
p.obj_clas_code as Prop_Obj_Clas_Code,
ps.prop_stts_abbr,nr.natr_rqst_abbr,
pi.pi_last_name, pi.pi_frst_name,
i.inst_shrt_name as inst_name,p.rqst_dol,p.rcom_awd_istr,ai.awd_istr_txt, ai.awd_istr_abbr,ai.awd_istr_abbr as rcom_istr_abbr,
p.prop_titl_txt,og.dir_div_abbr,id=identity(18), 0 as 'seq'
into #myProp
FROM (select distinct prop_id from #myBSplit mbs) as prop
JOIN csd.prop p ON p.prop_id = prop.prop_id
JOIN  #myBSplit mbs on mbs.prop_id=p.prop_id
JOIN csd.awd_istr ai on p.rcom_awd_istr = ai.awd_istr_code
JOIN csd.pgm_ele  as pe ON pe.pgm_ele_code=p.pgm_ele_code
JOIN csd.prop_stts ps on ps.prop_stts_code=p.prop_stts_code
JOIN csd.natr_rqst nr on nr.natr_rqst_code = p.natr_rqst_code
JOIN csd.pi_vw pi on p.pi_id=pi.pi_id
JOIN csd.inst as i on i.inst_id=p.inst_id
JOIN csd.org as og on og.org_code=p.org_code

SELECT getdate() as run_date,mp.nsf_rcvd_date,mp.dd_rcom_date,mp.Prop_Org_Code,
b.org_code as Budg_Org_Code,mp.PEC_prop,mbs.PEC_bdg_splt,
mp.Prop_Pm_ibm_logn_id,b.pm_ibm_logn_id as Budg_Pm_ibm_logn_id, mp.Prop_Obj_Clas_Code,b.obj_clas_code as Budg_Obj_Clas_Code,
mp.pgm_annc_id,mp.prop_stts_abbr,mp.natr_rqst_abbr,
 mp.dir_div_abbr,mbs.prop_id, mp.ILN,mp.lead_id,b.awd_id, mp.prop_titl_txt,
mbs.splt_id,mbs.budg_yr,b.budg_splt_tot_dol,mp.rqst_dol,
(SELECT MAX( CASE pa.seq WHEN 0 THEN rtrim(pa.prop_atr_code) END)+
MAX( CASE pa.seq WHEN 1 THEN ' '+rtrim(pa.prop_atr_code) END)+
MAX( CASE pa.seq WHEN 2 THEN ' '+rtrim(pa.prop_atr_code) END)+
MAX( CASE pa.seq WHEN 3 THEN ' '+rtrim(pa.prop_atr_code) END)+
MAX( CASE pa.seq WHEN 4 THEN ' '+rtrim(pa.prop_atr_code) END)+
MAX( CASE pa.seq WHEN 5 THEN ' '+rtrim(pa.prop_atr_code) END)+
MAX( CASE pa.seq WHEN 6 THEN ' '+rtrim(pa.prop_atr_code) END)+
MAX( CASE pa.seq WHEN 7 THEN ' '+rtrim(pa.prop_atr_code) END)+
MAX( CASE pa.seq WHEN 8 THEN ' '+rtrim(pa.prop_atr_code) END)+
MAX( CASE pa.seq WHEN 9 THEN ' '+rtrim(pa.prop_atr_code) END)+
MAX( CASE pa.seq WHEN 10 THEN ' '+rtrim(pa.prop_atr_code) END)
FROM #myPropPRCs pa WHERE pa.prop_id = mbs.prop_id) AS 'Prop PRCs',
(SELECT MAX( CASE bp.seq WHEN 0 THEN rtrim(bp.pgm_ref_code) END)+
MAX( CASE bp.seq WHEN 1 THEN ' ' + rtrim(bp.pgm_ref_code) END)+
MAX( CASE bp.seq WHEN 2 THEN ' ' + rtrim(bp.pgm_ref_code) END)+
MAX( CASE bp.seq WHEN 3 THEN ' ' + rtrim(bp.pgm_ref_code) END)+
MAX( CASE bp.seq WHEN 4 THEN ' ' + rtrim(bp.pgm_ref_code) END)+
MAX( CASE bp.seq WHEN 5 THEN ' ' + rtrim(bp.pgm_ref_code) END)+
MAX( CASE bp.seq WHEN 6 THEN ' ' + rtrim(bp.pgm_ref_code) END)+
MAX( CASE bp.seq WHEN 7 THEN ' ' + rtrim(bp.pgm_ref_code) END)+
MAX( CASE bp.seq WHEN 8 THEN ' ' + rtrim(bp.pgm_ref_code) END)+
MAX( CASE bp.seq WHEN 9 THEN ' ' + rtrim(bp.pgm_ref_code) END)+
MAX( CASE bp.seq WHEN 10 THEN ' ' + rtrim(bp.pgm_ref_code) END)
FROM #myBudgPRCs bp WHERE bp.prop_id=mbs.prop_id and bp.budg_yr=mbs.budg_yr and bp.splt_id = mbs.splt_id 
group by bp.prop_id, bp.budg_yr, bp.splt_id) AS 'Budg PRCs',
mp.pi_last_name, mp.pi_frst_name,mp.inst_name,b.last_updt_user, b.last_updt_tmsp,ast.awd_stts_abbr, ast.awd_stts_txt
, cs.cgi_stts_txt,mp.awd_istr_txt, mp.awd_istr_abbr,mp.rcom_istr_abbr 
FROM #myBSplit mbs
JOIN csd.budg_splt b on mbs.prop_id=b.prop_id and mbs.budg_yr=b.budg_yr and mbs.splt_id=b.splt_id
LEFT JOIN csd.awd a on b.awd_id=a.awd_id
LEFT JOIN csd.awd_stts ast on a.awd_stts_code=ast.awd_stts_code
LEFT JOIN csd.cgi c on  b.awd_id=c.awd_id and b.budg_yr= c.cgi_yr
LEFT JOIN csd.cgi_stts cs on c.cgi_stts_code=cs.cgi_stts_code
JOIN #myProp mp on mp.prop_id=mbs.prop_id

-- Awd Tracker
