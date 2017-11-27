-- what can you learn about proposal 1234567?
SELECT * FROM csd.prop WHERE prop_id = '1234567'

SELECT sc.colid, sc.name, t.name, sc.length 
FROM sysobjects so 
JOIN syscolumns sc ON sc.id = so.id 
JOIN systypes t on t.usertype = sc.usertype
WHERE so.name like 'pi_vw' -- table name
ORDER BY colid

select pi_id,pi_last_name,pi_frst_name,pi_mid_init,
prim_addr_flag,pi_emai_addr,
inst_id,pi_dept_name,pi_degr_yr,pi_acad_degr_txt,
pi_str_addr,pi_str_addr_addl,cty_name,st_code,ctry_code,zip_code,pi_phon_num,pi_fax_num,
pi_gend_code,pi_ctzn_code,pi_rno_code,pi_ethn_code,pi_hdcp_flag,pi_hdcp_othr_txt,
pi_actv_stts_code,oth_fed_proj_flag,
last_updt_tmsp,last_updt_pgm,last_updt_user
--pi_ssn,
--orc_id
from csd.pi_vw pi
where pi_last_name = 'Alphaman' and pi_frst_name = 'Alan'

select p.* from csd.pi_vw pi
join csd.prop p ON p.pi_id = pi.pi_id
where pi_last_name = 'Alphaman' and pi_frst_name = 'Alan'


-- who was the PI and what was the institution?
SELECT pi.pi_last_name, pi.pi_frst_name, inst.inst_shrt_name, p.*
FROM csd.prop p
JOIN csd.pi_vw pi ON pi.pi_id = p.pi_id
JOIN csd.inst inst ON inst.inst_id = p.inst_id
WHERE prop_id = '1234567'

-- Show your awards with the most recent action
SELECT TOP 10 * FROM csd.awd WHERE pm_ibm_logn_id = 'jsnoeyi' 
ORDER BY last_updt_tmsp DESC

-- Show upcoming panels  Sybase function getdate() return the current date/time.
SELECT TOP 10 * FROM csd.panl WHERE panl_bgn_date > getdate()

-- what are the possible review types and review status codes?
SELECT * FROM csd.rev_type
SELECT * FROM csd.rev_stts

-- Who were the top 20 reviewers in the number of FY17 panels for CISE?
SELECT TOP 20 r.revr_last_name, r.revr_frst_name, inst.inst_shrt_name, count(panl.panl_id) as nProp
FROM csd.panl panl
JOIN csd.panl_revr pr ON pr.panl_id = panl.panl_id
JOIN csd.revr r ON r.revr_id = pr.revr_id 
LEFT JOIN csd.inst inst ON inst.inst_id = r.inst_id
WHERE panl.org_code like '05%' AND panl.panl_id like 'P17%'
GROUP BY r.revr_last_name, r.revr_frst_name, inst.inst_shrt_name
ORDER By nProp DESC, revr_last_name, revr_frst_name

-- Ditto, but get reviewers into temporary table
SELECT TOP 20 pr.revr_id, count(panl.panl_id) as nProp
INTO #myPanlCounts
FROM csd.panl panl
JOIN csd.panl_revr pr ON pr.panl_id = panl.panl_id
WHERE panl.org_code like '05%' AND panl.panl_id like 'P17%'
GROUP BY pr.revr_id
ORDER BY nProp DESC
-- then identify them
SELECT r.revr_last_name, r.revr_frst_name, inst.inst_shrt_name, nProp
FROM #myPanlCounts pc
JOIN csd.revr r ON r.revr_id = pc.revr_id 
LEFT JOIN csd.inst inst ON inst.inst_id = r.inst_id
ORDER By nProp DESC, revr_last_name, revr_frst_name
DROP TABLE #myPanlCounts

select top 10 * from csd.rev_prop where prop_id like '17%'

-- How many reviewers reviewed for CISE in FY17?
SELECT count(DISTINCT revr_id) FROM csd.prop p 
JOIN csd.rev_prop rp ON p.prop_id = rp.prop_id
WHERE p.org_code LIKE '05%' AND p.prop_id LIKE '17%'

-- How many reviews for CISE in FY17?
SELECT count(revr_id) FROM csd.prop p 
JOIN csd.rev_prop rp ON p.prop_id = rp.prop_id
WHERE p.org_code LIKE '05%' AND p.prop_id LIKE '17%'



-- Of the reviewers for CISE, who are top 20 in number of FY17 reviews for NSF?
SELECT TOP 20 r.revr_last_name, r.revr_frst_name,inst.inst_shrt_name, count(rp2.prop_id) as nProp
FROM (SELECT DISTINCT revr_id FROM csd.prop p 
JOIN csd.rev_prop rp ON p.prop_id = rp.prop_id
WHERE p.org_code LIKE '05%' AND p.prop_id LIKE '17%') ciserevr
JOIN csd.rev_prop rp2 ON rp2.revr_id = ciserevr.revr_id AND rp2.prop_id LIKE '17%'
JOIN csd.revr r ON r.revr_id = ciserevr.revr_id
LEFT JOIN csd.inst inst ON inst.inst_id = r.inst_id
GROUP BY r.revr_last_name, r.revr_frst_name,inst.inst_shrt_name
ORDER BY nProp DESC, r.revr_last_name, r.revr_frst_name

SELECT count(rp2.prop_id)
FROM (SELECT DISTINCT revr_id FROM csd.prop p 
JOIN csd.rev_prop rp ON p.prop_id = rp.prop_id
WHERE p.org_code LIKE '05%' AND p.prop_id LIKE '17%') ciserevr
JOIN csd.rev_prop rp2 ON rp2.revr_id = ciserevr.revr_id AND rp2.prop_id LIKE '17%'
--JOIN csd.revr r ON r.revr_id = ciserevr.revr_id
--LEFT JOIN csd.inst inst ON inst.inst_id = r.inst_id
--GROUP BY r.revr_last_name, r.revr_frst_name,inst.inst_shrt_name
--ORDER BY nProp DESC, r.revr_last_name, r.revr_frst_name


-- Who were the top 20 reviewers in number of proposals reviewed for NSF in FY17?
SELECT TOP 20 rp.revr_id, count(rp.prop_id) as nProp
--INTO #myTopRevr
FROM csd.rev_prop rp 
WHERE rp.prop_id LIKE '17%'--AND rp.rev_stts_code <> 'C' AND rp.rev_rtrn_date is not null
GROUP BY rp.revr_id
ORDER BY nProp DESC

select top 100 * from csd.rev_prop

SELECT r.revr_last_name, r.revr_frst_name, inst.inst_shrt_name, t.nProp
FROM #myTopRevr t
JOIN csd.revr r ON r.revr_id = t.revr_id 
LEFT JOIN csd.inst inst ON inst.inst_id = r.inst_id
ORDER BY t.nProp DESC, r.revr_last_name, r.revr_frst_name
DROP TABLE #myTopRevr

-- For those who reviewed for CISE, who were top 20 in number of reviews for NSF?
SELECT TOP 20 r.revr_last_name, r.revr_frst_name, inst.inst_shrt_name, count(rp.prop_id) as nProp
FROM csd.rev_prop rp 
JOIN csd.revr r ON r.revr_id = rp.revr_id 
WHERE rp.org_code like '05%' AND rp.rev_stts_code <> 'C' AND rp.rev_rtrn_date is not null
ORDER BY nProp DESC

-- What is the assumption behind this analysis? (unique reviewer ids, props vs projects)
SELECT TOP 20 r.revr_last_name, r.revr_frst_name, inst.inst_shrt_name, count(rp.prop_id) as nProp
FROM csd.rev_prop rp 
JOIN csd.revr r ON r.revr_id = rp.revr_id 
WHERE rp.org_code like '05%' AND rp.rev_stts_code <> 'C' AND rp.rev_rtrn_date is not null
ORDER BY nProp DESC

-- Count CISE proposals in FY17 BY program
SELECT p.pgm_ele_code, pec.pgm_ele_name, count(p.prop_id) AS nProp  
FROM csd.prop p
JOIN csd.pgm_ele pec ON pec.pgm_ele_code = p.pgm_ele_code
WHERE prop_id LIKE '17%' and org_code LIKE '05%'
GROUP BY p.pgm_ele_code, pec.pgm_ele_name
ORDER BY nProp DESC

-- Count CISE projects in FY17 BY program
SELECT p.pgm_ele_code, pec.pgm_ele_name, count(p.prop_id) AS nProp  
FROM csd.prop p
JOIN csd.pgm_ele pec ON pec.pgm_ele_code = p.pgm_ele_code
WHERE prop_id LIKE '17%' and org_code LIKE '05%' AND (lead_prop_id is null OR lead_prop_id=prop_id)
GROUP BY p.pgm_ele_code, pec.pgm_ele_name
ORDER BY nProp DESC

-- Count CISE projects in FY17 BY program
SELECT p.pgm_ele_code, pec.pgm_ele_name, count(p.prop_id) AS nProp  
FROM csd.prop p
JOIN csd.pgm_ele pec ON pec.pgm_ele_code = p.pgm_ele_code
WHERE prop_id LIKE '17%' and org_code LIKE '05%' AND isnull(lead_prop_id,prop_id)=prop_id
GROUP BY p.pgm_ele_code, pec.pgm_ele_name
ORDER BY nProp DESC


-- who are the top 20 PDs at approving reports in FY17?
--Reports query:
declare @startdate datetime, @enddate datetime
select @startdate = '10/1/2016', @enddate = '10/1/2017'
SELECT r.aprv_user_id, COUNT(r.awd_id) as Nrpt, AVG(1.0*datediff(day,r.SUB_DATE, r.APRV_DATE)) AvgAprvTime, STDEV(datediff(day,r.SUB_DATE, r.APRV_DATE)) StDevAprvTime
FROM csd.pr_cntl r
WHERE (r.APRV_DATE BETWEEN @startdate AND @enddate)
AND r.aprv_user_id is not null AND rpt_type <> 'O'
GROUP BY r.aprv_user_id

-- What are the top 20 PDs in panels managed for FY17?
SELECT po.p.*
FROM (SELECT TOP 20 pm_logn_id,COUNT(panl_id) as Npanl
      FROM csd.panl pn 
      WHERE panl_id like 'P17%'
      GROUP BY pm_logn_id) p
JOIN csd.po_vw po ON po.po_ibm_logn_id = left(pn.pm_logn_id,7)

--Projects/proposals:
declare @startdate datetime, @enddate datetime
select @startdate = '10/1/2016', @enddate = '10/1/2017'
SELECT p.*, natr_rqst_abbr, obj_clas_name, pgm_ele_name 
    FROM (SELECT  p.pm_ibm_logn_id, left(p.org_code,4) as Org, p.pgm_annc_id,p.pgm_ele_code, p.natr_rqst_code, p.obj_clas_code, prop_stts_abbr, 
    COUNT(p.prop_id) as Nprop, SUM(CASE WHEN lead_prop_id<>p.prop_id THEN 0 ELSE 1 END) as Nproj, Sum(p.rqst_dol) as SumRqstDol
    FROM csd.prop p
    JOIN csd.prop_stts ps ON ps.prop_stts_code = p.prop_stts_code
    WHERE (p.dd_rcom_date BETWEEN @startdate AND @enddate)
    AND obj_clas_code LIKE '4%'
    GROUP BY pm_ibm_logn_id, left(p.org_code,4), p.pgm_annc_id, p.pgm_ele_code, natr_rqst_code, p.obj_clas_code, prop_stts_abbr) p
JOIN csd.natr_rqst nr ON nr.natr_rqst_code = p.natr_rqst_code
JOIN csd.obj_clas oc ON oc.obj_clas_code = p.obj_clas_code
JOIN csd.pgm_ele pec ON pec.pgm_ele_code = p.pgm_ele_code

--Budget splits query:
SELECT b.pm_ibm_logn_id, left(b.org_code,4) as Org, b.pgm_ele_code, b.obj_clas_code, 
a.pm_ibm_logn_id, left(a.org_code,4) as aOrg, a.pgm_ele_code as aPEC,
COUNT(DISTINCT prop_id) as Nsplit,
--SUM(CASE WHEN (last_updt_tmsp BETWEEN @startdate AND @enddate)  THEN 1 ELSE 0 END) AS fyAction,
SUM(budg_splt_tot_dol) AS budg_tot
FROM csd.budg_splt b
LEFT JOIN csd.awd a ON a.awd_id = b.awd_id
WHERE budg_yr = 2017 AND obj_clas_code LIKE '4%' 
GROUP BY b.pm_ibm_logn_id, left(b.org_code,4), b.pgm_ele_code, b.obj_clas_code, 
a.pm_ibm_logn_id, left(a.org_code,4), a.pgm_ele_code

--Panels query:
declare @startdate datetime, @enddate datetime
select @startdate = '10/1/2016', @enddate = '10/1/2017'

SELECT pm_logn_id,left(org_code,4) as Org,pgm_ele_code, COUNT(panl_id) as Npanl
FROM csd.panl pn 
WHERE (pn.panl_bgn_date BETWEEN @startdate AND @enddate)
GROUP BY pm_logn_id,left(org_code,4),pgm_ele_code

--- Using systables to determine all rptdb tables with similarly-named fields.
-- Which are lan_ids and which are ibm_logn_ids?
select so.name, sc.name, t.name, sc.length 
from syscolumns sc 
JOIN sysobjects so ON sc.id = so.id
JOIN systypes t on t.usertype = sc.usertype
where sc.name like '%p%logn%'
order by so.name, sc.name
