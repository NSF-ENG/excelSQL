-- what can you learn about proposal 1743463
SELECT * FROM csd.prop WHERE prop_id = '1743463'

-- who was the PI and what was the institution?
SELECT pi.pi_last_name, pi.pi_frst_name, inst.inst_shrt_name, p.*
FROM csd.prop p
JOIN csd.pi_vw pi ON pi.pi_id = p.pi_id
LEFT JOIN csd.inst inst ON inst.inst_id = p.inst_id
WHERE prop_id = '1743463'

--find all about all proposals FROM Alan Alphaman
SELECT p.* FROM csd.pi_vw pi
JOIN csd.prop p ON p.pi_id = pi.pi_id
WHERE pi_last_name = 'Alphaman' and pi_frst_name = 'Alan'

-- does Alan serve AS a co-PI?
SELECT a.* FROM csd.pi_vw pi
JOIN csd.addl_pi_invl a ON a.pi_id = pi.pi_id
WHERE pi_last_name = 'Alphaman' and pi_frst_name = 'Alan'

--find all about awards to Alan Alphaman: BY pi or BY proposal
SELECT a.* FROM csd.pi_vw pi
JOIN csd.awd a ON a.pi_id = pi.pi_id
WHERE pi_last_name = 'Alphaman' and pi_frst_name = 'Alan'

SELECT a.* FROM csd.pi_vw pi
JOIN csd.prop p ON p.pi_id = pi.pi_id
JOIN csd.awd a ON a.awd_id = p.prop_id
WHERE pi_last_name = 'Alphaman' and pi_frst_name = 'Alan'

--why can't we find all about pis?  
SELECT * FROM csd.pi_vw pi
WHERE pi_last_name = 'Alphaman' and pi_frst_name = 'Alan'

--ask the database about the fields
-- table, column names & types 
SELECT so.name AS tablename, sc.colid, sc.name AS fieldname, t.name AS typename, sc.length 
FROM sysobjects so 
JOIN syscolumns sc ON sc.id = so.id 
JOIN systypes t ON t.usertype = sc.usertype
WHERE so.name LIKE 'pi_vw' and sc.name LIKE '%' -- table name, field name
ORDER BY so.name, colid

-- with protection information, too. Note: some DIS queries use pi; we need to use pi_vw.
SELECT so.name AS tablename, so.crdate, sc.colid, sc.name AS fieldname, t.name AS typename, sc.length, su.name AS username, sp.uid, sp.action, sp.protecttype, 
inttohex(sp.columns) AS col
FROM sysobjects so 
JOIN syscolumns sc ON sc.id = so.id
JOIN systypes t ON t.usertype = sc.usertype
LEFT JOIN sysprotects sp ON sp.id = sc.id
LEFT JOIN sysusers su ON su.uid = sp.uid
WHERE so.name LIKE 'pi%' and sc.name LIKE '%' -- table name, field name
AND (su.name is null OR su.name = 'public' OR su.name LIKE 'ccfuser')
ORDER BY so.name, colid, su.name

-- Here are the fields we can look at with in pi_vw
SELECT pi_id,pi_last_name,pi_frst_name,pi_mid_init,
prim_addr_flag,pi_emai_addr,
inst_id,pi_dept_name,pi_degr_yr,pi_acad_degr_txt,
pi_str_addr,pi_str_addr_addl,cty_name,st_code,ctry_code,zip_code,pi_phon_num,pi_fax_num,
pi_gend_code,pi_ctzn_code,pi_rno_code,pi_ethn_code,pi_hdcp_flag,pi_hdcp_othr_txt,
pi_actv_stts_code,oth_fed_proj_flag,
last_updt_tmsp,last_updt_pgm,last_updt_user
--pi_ssn,
--orc_id
FROM csd.pi_vw pi
WHERE pi_last_name = 'Alphaman' and pi_frst_name = 'Alan'

--- Using systables to determine all rptdb tables with similarly-named fields.
-- Which are lan_ids and which are ibm_logn_ids?
SELECT so.name AS tablename, sc.colid, sc.name AS fieldname, t.name AS typename, sc.length 
FROM sysobjects so 
JOIN syscolumns sc ON sc.id = so.id 
JOIN systypes t ON t.usertype = sc.usertype
WHERE so.name LIKE '%' and sc.name LIKE '%p%log%' -- table name, field name
ORDER BY so.name, colid


-- Show your awards with the most recent action
SELECT TOP 10 * FROM csd.awd WHERE pm_ibm_logn_id = 'jsnoeyi' 
ORDER BY last_updt_tmsp DESC

-- Count proposals (projects) you handled in FY17
SELECT ps.prop_stts_abbr, nr.natr_rqst_abbr, count(p.prop_id) AS nProp
FROM csd.prop p
JOIN csd.prop_stts ps ON ps.prop_stts_code = p.prop_stts_code
JOIN csd.natr_rqst nr ON nr.natr_rqst_code = p.natr_rqst_code
WHERE p.prop_id LIKE '17%' AND p.pm_ibm_logn_id = 'jsnoeyi' 
GROUP BY ps.prop_stts_abbr, nr.natr_rqst_abbr
ORDER BY nProp DESC

-- Count projects you handled in FY17: using OR
SELECT ps.prop_stts_abbr, nr.natr_rqst_abbr, count(p.prop_id) AS nProp
FROM csd.prop p
JOIN csd.prop_stts ps ON ps.prop_stts_code = p.prop_stts_code
JOIN csd.natr_rqst nr ON nr.natr_rqst_code = p.natr_rqst_code
WHERE p.prop_id LIKE '17%' AND p.pm_ibm_logn_id = 'jsnoeyi' 
AND (p.lead_prop_id is NULL OR p.lead_prop_id = p.prop_id)
GROUP BY ps.prop_stts_abbr, nr.natr_rqst_abbr
ORDER BY nProp DESC

-- Count projects you handled in FY17 using isnull.
SELECT ps.prop_stts_abbr, nr.natr_rqst_abbr, count(p.prop_id) AS nProp
FROM csd.prop p
JOIN csd.prop_stts ps ON ps.prop_stts_code = p.prop_stts_code
JOIN csd.natr_rqst nr ON nr.natr_rqst_code = p.natr_rqst_code
WHERE p.prop_id LIKE '17%' AND p.pm_ibm_logn_id = 'jsnoeyi' 
AND isnull(p.lead_prop_id,p.prop_id) = p.prop_id
GROUP BY ps.prop_stts_abbr, nr.natr_rqst_abbr
ORDER BY nProp DESC

-- who are the top 20 PDs at approving annual and final reports in FY17?
declare @startdate datetime, @enddate datetime
SELECT @startdate = '10/1/2016', @enddate = '10/1/2017'
SELECT TOP 100 r.aprv_user_id, COUNT(r.awd_id) AS Nrpt, AVG(1.0*datediff(day,r.SUB_DATE, r.APRV_DATE)) AvgAprvTime, STDEV(datediff(day,r.SUB_DATE, r.APRV_DATE)) StDevAprvTime
FROM csd.pr_cntl r
WHERE (r.APRV_DATE BETWEEN @startdate AND @enddate)
AND r.aprv_user_id is not null AND rpt_type <> 'O'
GROUP BY r.aprv_user_id
ORDER BY Nrpt DESC


-- Count CISE proposals (projects) in FY17 BY program
SELECT p.pgm_ele_code, pec.pgm_ele_name, count(p.prop_id) AS nProp  
FROM csd.prop p
JOIN csd.pgm_ele pec ON pec.pgm_ele_code = p.pgm_ele_code
WHERE prop_id LIKE '17%' and org_code LIKE '05%' 
--AND (lead_prop_id is null OR lead_prop_id=prop_id)
--AND isnull(lead_prop_id,prop_id)=prop_id
GROUP BY p.pgm_ele_code, pec.pgm_ele_name
ORDER BY nProp DESC

-- Show upcoming panels  Sybase function getdate() return the current date/time.
SELECT TOP 10 * FROM csd.panl WHERE panl_bgn_date > getdate()

-- How many panels not in CISE reviewed how many CISE projects?
SELECT pn.org_code, count(DISTINCT pn.panl_id) as nPanl, count(pp.prop_id)
FROM csd.prop p 
JOIN csd.panl_prop pp ON pp.prop_id = p.prop_id
JOIN csd.panl pn ON pn.panl_id = pp.panl_id
WHERE p.org_code LIKE '05%' AND p.prop_id LIKE '17%' --AND pn.org_code NOT LIKE '05%'
AND isnull(p.lead_prop_id,p.prop_id) = p.prop_id
--AND p.pgm_annc_id in ('NSF 16-578','NSF 16-579','NSF 16-581') -- core
--AND p.pgm_annc_id BETWEEN 'NSF 16-578' AND 'NSF 16-581' -- core+satc
--AND NOT p.pgm_annc_id BETWEEN 'NSF 16-578' AND 'NSF 16-581' -- core+satc
GROUP BY pn.org_code
ORDER BY pn.org_code

-- Who were the top 20 reviewers BY the number of FY17 panels for CISE?
SELECT TOP 20 r.revr_last_name, r.revr_frst_name, inst.inst_shrt_name, count(panl.panl_id) AS nProp
FROM csd.panl panl
JOIN csd.panl_revr pr ON pr.panl_id = panl.panl_id
JOIN csd.revr r ON r.revr_id = pr.revr_id 
LEFT JOIN csd.inst inst ON inst.inst_id = r.inst_id
WHERE panl.org_code LIKE '05%' AND panl.panl_id LIKE 'P17%'
GROUP BY r.revr_last_name, r.revr_frst_name, inst.inst_shrt_name
ORDER By nProp DESC, revr_last_name, revr_frst_name

-- Ditto, but get reviewers into temporary table: explain the difference.
SELECT TOP 20 pr.revr_id, count(panl.panl_id) AS nProp
INTO #myPanlCounts
FROM csd.panl panl
JOIN csd.panl_revr pr ON pr.panl_id = panl.panl_id
WHERE panl.org_code LIKE '05%' AND panl.panl_id LIKE 'P17%'
GROUP BY pr.revr_id
ORDER BY nProp DESC
-- then identify them
SELECT r.revr_last_name, r.revr_frst_name, inst.inst_shrt_name, nProp
FROM #myPanlCounts pc
JOIN csd.revr r ON r.revr_id = pc.revr_id 
LEFT JOIN csd.inst inst ON inst.inst_id = r.inst_id
ORDER By nProp DESC, revr_last_name, revr_frst_name
DROP TABLE #myPanlCounts


-- linking proposals to their reviews
SELECT top 10 * FROM csd.rev_prop WHERE prop_id LIKE '17%'

-- None of alan alphaman's proposals have reviews
SELECT pi.pi_id, p.lead_prop_id, p.prop_id, rp.* FROM csd.pi_vw pi
JOIN csd.prop p ON p.pi_id = pi.pi_id
LEFT JOIN csd.rev_prop rp ON rp.prop_id = p.prop_id
WHERE pi_last_name = 'Alphaman' and pi_frst_name = 'Alan'


-- what are the possible review types and review status codes?
SELECT * FROM csd.rev_type
SELECT * FROM csd.rev_stts

-- How many reviews and reviewers for CISE in FY17 ?
SELECT count(revr_id) AS nRev, count(DISTINCT revr_id) AS nRevr 
FROM csd.prop p 
JOIN csd.rev_prop rp ON rp.prop_id = p.prop_id AND rp.rev_stts_code <> 'C' AND rp.rev_rtrn_date is not null
WHERE p.org_code LIKE '05%' AND p.prop_id LIKE '17%'
--AND p.pgm_annc_id in ('NSF 16-578','NSF 16-579','NSF 16-581') -- core
--AND p.pgm_annc_id BETWEEN 'NSF 16-578' AND 'NSF 16-581' -- core+satc
--AND NOT p.pgm_annc_id BETWEEN 'NSF 16-578' AND 'NSF 16-581' -- not core;satc

-- how many reviews returned with conflicts?
SELECT count(revr_id) AS nRev, count(DISTINCT revr_id) AS nRevr 
FROM csd.prop p 
JOIN csd.rev_prop rp ON rp.prop_id = p.prop_id AND rp.rev_stts_code = 'C' AND rp.rev_rtrn_date is not null
WHERE p.org_code LIKE '05%' AND p.prop_id LIKE '17%'



-- Of the reviewers for CISE, who are top 20 in number of FY17 reviews for NSF?
SELECT TOP 20 r.revr_last_name, r.revr_frst_name,inst.inst_shrt_name, count(rp2.prop_id) AS nProp
FROM (SELECT DISTINCT revr_id FROM csd.prop p 
JOIN csd.rev_prop rp ON p.prop_id = rp.prop_id
WHERE p.org_code LIKE '05%' AND p.prop_id LIKE '17%') ciserevr
JOIN csd.rev_prop rp2 ON rp2.revr_id = ciserevr.revr_id AND rp2.prop_id LIKE '17%' AND rp.rev_stts_code <> 'C' AND rp.rev_rtrn_date is not null
JOIN csd.revr r ON r.revr_id = ciserevr.revr_id
LEFT JOIN csd.inst inst ON inst.inst_id = r.inst_id
GROUP BY r.revr_last_name, r.revr_frst_name,inst.inst_shrt_name
ORDER BY nProp DESC, r.revr_last_name, r.revr_frst_name

SELECT count(rp2.prop_id)
FROM (SELECT DISTINCT revr_id FROM csd.prop p 
JOIN csd.rev_prop rp ON p.prop_id = rp.prop_id AND rp.rev_stts_code <> 'C' AND rp.rev_rtrn_date is not null
WHERE p.org_code LIKE '05%' AND p.prop_id LIKE '17%') ciserevr
JOIN csd.rev_prop rp2 ON rp2.revr_id = ciserevr.revr_id AND rp2.prop_id LIKE '17%'
--JOIN csd.revr r ON r.revr_id = ciserevr.revr_id
--LEFT JOIN csd.inst inst ON inst.inst_id = r.inst_id
--GROUP BY r.revr_last_name, r.revr_frst_name,inst.inst_shrt_name
--ORDER BY nProp DESC, r.revr_last_name, r.revr_frst_name

-- Who were the top 100 reviewers in number of proposals reviewed for NSF in FY17?
SELECT TOP 100 rp.revr_id, count(rp.prop_id) AS nProp
INTO #myTopRevr
FROM csd.rev_prop rp 
WHERE rp.prop_id LIKE '17%' AND rp.rev_stts_code <> 'C' AND rp.rev_rtrn_date is not null
GROUP BY rp.revr_id
ORDER BY nProp DESC

SELECT r.revr_last_name, r.revr_frst_name, inst.inst_shrt_name, t.nProp
FROM #myTopRevr t
JOIN csd.revr r ON r.revr_id = t.revr_id 
LEFT JOIN csd.inst inst ON inst.inst_id = r.inst_id
ORDER BY t.nProp DESC, r.revr_last_name, r.revr_frst_name
DROP TABLE #myTopRevr

-- What are the assumptions behind this analysis? (unique reviewer ids, props vs projects)
SELECT TOP 20 r.revr_last_name, r.revr_frst_name, inst.inst_shrt_name, count(rp.prop_id) AS nProp
FROM csd.rev_prop rp 
JOIN csd.revr r ON r.revr_id = rp.revr_id 
WHERE rp.org_code LIKE '05%' AND rp.rev_stts_code <> 'C' AND rp.rev_rtrn_date is not null
ORDER BY nProp DESC




-- What are the top 20 PDs in panels managed for FY17?
SELECT po.p.*
FROM (SELECT TOP 20 pm_logn_id,COUNT(panl_id) AS Npanl
      FROM csd.panl pn 
      WHERE panl_id LIKE 'P17%'
      GROUP BY pm_logn_id) p
JOIN csd.po_vw po ON po.po_ibm_logn_id = left(pn.pm_logn_id,7)

--Projects/proposals:
declare @startdate datetime, @enddate datetime
SELECT @startdate = '10/1/2016', @enddate = '10/1/2017'
SELECT p.*, natr_rqst_abbr, obj_clas_name, pgm_ele_name 
    FROM (SELECT  p.pm_ibm_logn_id, left(p.org_code,4) AS Org, p.pgm_annc_id,p.pgm_ele_code, p.natr_rqst_code, p.obj_clas_code, prop_stts_abbr, 
    COUNT(p.prop_id) AS Nprop, SUM(CASE WHEN lead_prop_id<>p.prop_id THEN 0 ELSE 1 END) AS Nproj, Sum(p.rqst_dol) AS SumRqstDol
    FROM csd.prop p
    JOIN csd.prop_stts ps ON ps.prop_stts_code = p.prop_stts_code
    WHERE (p.dd_rcom_date BETWEEN @startdate AND @enddate)
    AND obj_clas_code LIKE '4%'
    GROUP BY pm_ibm_logn_id, left(p.org_code,4), p.pgm_annc_id, p.pgm_ele_code, natr_rqst_code, p.obj_clas_code, prop_stts_abbr) p
JOIN csd.natr_rqst nr ON nr.natr_rqst_code = p.natr_rqst_code
JOIN csd.obj_clas oc ON oc.obj_clas_code = p.obj_clas_code
JOIN csd.pgm_ele pec ON pec.pgm_ele_code = p.pgm_ele_code

--Budget splits query:
SELECT b.pm_ibm_logn_id, left(b.org_code,4) AS Org, b.pgm_ele_code, b.obj_clas_code, 
a.pm_ibm_logn_id, left(a.org_code,4) AS aOrg, a.pgm_ele_code AS aPEC,
COUNT(DISTINCT prop_id) AS Nsplit,
--SUM(CASE WHEN (last_updt_tmsp BETWEEN @startdate AND @enddate)  THEN 1 ELSE 0 END) AS fyAction,
SUM(budg_splt_tot_dol) AS budg_tot
FROM csd.budg_splt b
LEFT JOIN csd.awd a ON a.awd_id = b.awd_id
WHERE budg_yr = 2017 AND obj_clas_code LIKE '4%' 
GROUP BY b.pm_ibm_logn_id, left(b.org_code,4), b.pgm_ele_code, b.obj_clas_code, 
a.pm_ibm_logn_id, left(a.org_code,4), a.pgm_ele_code

--Panels query:
declare @startdate datetime, @enddate datetime
SELECT @startdate = '10/1/2016', @enddate = '10/1/2017'

SELECT pm_logn_id,left(org_code,4) AS Org,pgm_ele_code, COUNT(panl_id) AS Npanl
FROM csd.panl pn 
WHERE (pn.panl_bgn_date BETWEEN @startdate AND @enddate)
GROUP BY pm_logn_id,left(org_code,4),pgm_ele_code


