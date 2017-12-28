-- RoboRA sql code   Jack Snoeyink Nov 2017
--RoboRA queries pull everything from the database that might go into an RA 
-- to support creation of RA drafts from RAtemplates in Word mail merge.
-- It separately pulling text fields to support creation of an indexed pdf.  
-- (We don't pull text into drafts, to avoid duplicating material already in the jacket.)
--In addition it makes several sheets for checking coding of both declines and awards, 
-- and makes budget numbers available. 

-- Commments with [name and ]name delimit code that will be copied into the hidden SQL worksheet of RoboRA.xlsm
-- You can simply copy the entire thing into the clipboard and run "saveSQLcode" on that sheet.
-- All other code is for testing.  Some comments are for developing long queries in pieces.
--
-- When SQL has run-time parameters, either break into static strings to use in VBA as str1 & param & str2 
--    or use @varnames that will be declared before your string
--    "declare @varname char(7), @datename datetime" & vbNewline & "SELECT @varname = " & param1 &" @datename = "& Now

-- get lists of relevant prop_ids in many ways
-- This will mostly be generated on the fly. 
--[RA_pidRAt
--DROP TABLE #myPidRAt
SET NOCOUNT ON 
CREATE TABLE #myPidRAt (lead_prop_id char(7) null, prop_id char(7), pi_id char(9), inst_id char(10), PO char(7), RAtemplate varchar(63))
--]RA_pidRAt
INSERT INTO #myPidRAt 
--[RA_pidRAtSelect
SELECT DISTINCT p.lead_prop_id, p.prop_id, p.pi_id, p.inst_id, p.pm_ibm_logn_id as PO, 
--]RA_pidRAtSelect
'' as RAtemplate 
FROM csd.prop p
JOIN csd.panl_prop pp ON p.prop_id = pp.prop_id
WHERE p.prop_stts_code LIKE '0[01289]' AND
pp.panl_id in ('p172027','p170288','p180207','p180208')
--pm_logn_id = 'jsnoeyi'


-- Every other query will begin with this: 
--  Add collabs not already there, then make props, leads, and get RA update time
--[RA_leads
-- Needs: #myPidRAt
-- DROP TABLE #myProp, #myLead, #myRA
CREATE INDEX myPid_idx ON #myPidRAt(prop_id)
INSERT INTO #myPidRAt SELECT DISTINCT p.lead_prop_id, p.prop_id, p.pi_id, p.inst_id, p.pm_ibm_logn_id as PO, pid.RAtemplate 
FROM #myPidRAt pid
JOIN csd.prop p ON p.lead_prop_id = pid.lead_prop_id
WHERE pid.lead_prop_id is not NULL  
    AND NOT EXISTS (SELECT * FROM #myPidRAt px WHERE px.prop_id = p.prop_id AND px.RAtemplate = pid.RAtemplate)
--get props
--  prop.lead_prop_id, prop.pi_id, prop.inst_id, prop.pm_ibm_logn_id as PO, 
SELECT CASE WHEN p.lead_prop_id IS NULL THEN 'I' WHEN p.lead_prop_id <> p.prop_id THEN 'N' ELSE 'L' END AS ILN, 
  isnull(p.lead_prop_id,p.prop_id) AS lead, p.prop_id, p.pi_id, pi_last_name as L, pi_frst_name as F, 
  inst_shrt_name as I, inst.st_code, pi_emai_addr AS M, p.PO
INTO #myProp FROM (SELECT DISTINCT lead_prop_id, prop_id, pi_id, inst_id, PO FROM #myPidRAt) p
LEFT JOIN csd.pi_vw pi ON pi.pi_id = p.pi_id AND pi.prim_addr_flag='Y'
LEFT JOIN csd.inst inst ON inst.inst_id = p.inst_id
ORDER BY lead, ILN, p.prop_id
CREATE INDEX myProp_idx ON #myProp(prop_id)
--get leads
SELECT p.ILN, p.PO, p.lead, p.L
INTO #myLead FROM #myProp p WHERE p.ILN < 'M' 
CREATE INDEX myLead_idx ON #myLead(lead)
-- determine if we have an RA (doc_type_code '034')
SELECT lead, MAX(last_updt_tmsp) as RAupdate -- check text and eJupload for last RA
INTO #myRA
FROM (SELECT p.lead, ra.last_updt_tmsp
    FROM #myLead p
    JOIN csd.prop_rev_anly_vw ra ON ra.prop_id = p.lead
    UNION ALL SELECT p.lead, ej.last_updt_tmsp
    FROM #myLead p
    JOIN csd.ej_upld_doc_vw ej ON ej.prop_id = p.lead AND ej_doc_type_code = '034') d
GROUP BY lead
--]RA_leads
-- select * from #myProp select * from #myLead select * from #myRA

--[RA_PECglossary
--Needs: RA_lead(#myProp) 
SELECT p.pgm_ele_code AS PEC, pec.pgm_ele_name AS PEC_Description
FROM (SELECT DISTINCT pgm_ele_code FROM #myProp pid 
JOIN csd.prop p ON p.prop_id = pid.prop_id) p
JOIN FLflpdb.flp.pgm_ele pec ON pec.pgm_ele_code = p.pgm_ele_code
ORDER BY PEC
--]RA_PECglossary

-- check institutions of reviewers with conflicts
select top 200 pp.prop_id, panl_id, count(revr_id) as nConfl
from csd.panl_prop pp 
JOIN csd.rev_prop rp on rp.prop_id = pp.prop_id AND rp.rev_stts_code = 'C'
where panl_id like 'P17%'
group by pp.prop_id, panl_id
order by nConfl desc

--[RA_ckConfRevrInst
--Needs: RA_lead(#myLead)
SELECT r.revr_id, i.inst_shrt_name, rtrim(revr.revr_last_name) + ', ' + rtrim(revr.revr_frst_name) AS Panelist
FROM (SELECT DISTINCT rp.revr_id FROM #myLead p
     JOIN csd.rev_prop rp ON rp.prop_id = p.lead AND rp.rev_stts_code = 'C') as r
JOIN csd.revr revr ON revr.revr_id = r.revr_id
LEFT JOIN csd.inst i ON i.inst_id = revr.inst_id
ORDER BY i.inst_shrt_name, Panelist
--]RA_ckConfRevrInst

-- PRCs for props: all
--[RA_propPRCs
-- Needs: RA_lead(#myProp)
-- DROP TABLE #myPRCs, #myPRCdata
SELECT prc.*, id=identity(18), 0 as seq
INTO #myPRCdata
FROM (SELECT DISTINCT p.prop_id, pa.prop_atr_code
    FROM #myProp p
    JOIN csd.prop_atr pa ON pa.prop_id = p.prop_id AND pa.prop_atr_type_code = 'PRC') prc
ORDER BY prop_id, prop_atr_code
SELECT prop_id, MIN(id) as start INTO #myPRCStart FROM #myPRCdata GROUP BY prop_id
UPDATE #myPRCdata SET seq = id-M.start FROM #myPRCdata r, #myPRCStart M WHERE r.prop_id = M.prop_id
SELECT prop_id,convert(varchar(35),
    MAX( CASE pa.seq WHEN 0 THEN       pa.prop_atr_code ELSE '' END ) +
    MAX( CASE pa.seq WHEN 1 THEN ' ' + pa.prop_atr_code ELSE '' END ) +
    MAX( CASE pa.seq WHEN 2 THEN ' ' + pa.prop_atr_code ELSE '' END ) +
    MAX( CASE pa.seq WHEN 3 THEN ' ' + pa.prop_atr_code ELSE '' END ) +
    MAX( CASE pa.seq WHEN 4 THEN ' ' + pa.prop_atr_code ELSE '' END ) +
    MAX( CASE pa.seq WHEN 5 THEN ' ' + pa.prop_atr_code ELSE '' END ) +
    MAX( CASE pa.seq WHEN 6 THEN ' ' + pa.prop_atr_code ELSE '' END ) +
    MAX( CASE pa.seq WHEN 7 THEN ' ' + pa.prop_atr_code ELSE '' END ) +
    MAX( CASE pa.seq WHEN 8 THEN ' ' + pa.prop_atr_code ELSE '' END ) +
    MAX( CASE pa.seq WHEN 9 THEN ' ' + pa.prop_atr_code ELSE '' END ) +
    MAX( CASE pa.seq WHEN 10 THEN ' ' + pa.prop_atr_code ELSE '' END )) AS R
INTO #myPRCs
FROM #myPRCdata pa 
GROUP BY prop_id
DROP TABLE #myPRCStart
CREATE INDEX myPRCs_ix ON #myPRCs(prop_id)
--]RA_propPRCs
--select * from #myPRCs

---- all budget revisions DROP TABLE #myBudg
--SELECT p.prop_id, eb.revn_num, eb.budg_seq_yr, eb.budg_tot_dol, 
--sub_ctr_dol, frgn_trav_dol, pdoc_grnt_dol,part_dol, grad_pers_tot_cnt,
--sr_pers_cnt, sr_summ_mnth_cnt, sr_acad_mnth_cnt, sr_cal_mnth_cnt
--INTO #myBudg
--FROM #myProp p
--JOIN csd.eps_blip eb ON p.prop_id = eb.prop_id 
----AND NOT EXISTS (SELECT eb1.revn_num FROM csd.eps_blip eb1 WHERE eb.prop_id = eb1.prop_id AND eb.revn_num < eb1.revn_num)
--ORDER BY p.prop_id, eb.revn_num, eb.budg_seq_yr 
--CREATE INDEX myBudg_ix ON #myBudg(prop_id)

-- totals for project in last budget revision DROP TABLE #myPropBudg, #myPropInfo
--select * from #myPropBudg
--[RA_budg
-- Needs: RA_lead(#myProp)
--DROP TABLE #myBudg
SELECT p.lead, p.prop_id, eb.revn_num, eb.budg_seq_yr, eb.budg_tot_dol,
sub_ctr_dol, frgn_trav_dol, pdoc_grnt_dol, part_dol, grad_pers_tot_cnt,
sr_pers_cnt,sr_summ_mnth_cnt,sr_acad_mnth_cnt, sr_cal_mnth_cnt
INTO #myBudg FROM #myProp p
JOIN csd.eps_blip eb ON eb.prop_id = p.prop_id 
    AND NOT EXISTS (SELECT eb1.revn_num FROM csd.eps_blip eb1 WHERE eb.prop_id = eb1.prop_id AND eb.revn_num < eb1.revn_num)
ORDER BY prop_id, budg_seq_yr
CREATE INDEX myBudg_idx ON #myBudg(prop_id, budg_seq_yr)
--]RA_budg

-- # of projects whose last budget year is lastYr
--select lastYr, count(prop_id) as total FROM
--(select eb.prop_id, Max(budg_seq_yr) as lastYr
--  from csd.eps_blip eb 
--  Join csd.prop p ON p.prop_id = eb.prop_id AND (lead_prop_id is null OR lead_prop_id = eb.prop_id)
--  where eb.prop_id like '17%'
-- group by eb.prop_id) p
--group by lastYr
--order by lastYr desc
--

-- totals for project in last budget revision, per-proposal info
--select * from #myPropBudg
--[RA_prop
-- Needs: RA_lead(#myProp), RA_propPRCs(#myPRCs), RA_budg(#myBudg)
--DROP TABLE #myPropBudg, #myPropInfo
SELECT p.prop_id, eb.revn_num as RN, SUM(eb.budg_tot_dol) as T,
 nullif(SUM(sub_ctr_dol),0) AS sub_ctr_tot, nullif(SUM(frgn_trav_dol),0) AS frgn_trav_tot, nullif(SUM(pdoc_grnt_dol),0) AS pdoc_tot,nullif(SUM(part_dol),0) AS part_tot_dol, nullif(SUM(grad_pers_tot_cnt),0) AS grad_tot_cnt,
 nullif(SUM(sr_pers_cnt),0) AS sr_tot_cnt, nullif(SUM(sr_summ_mnth_cnt),0) AS sr_sumr_mnths, nullif(SUM(sr_acad_mnth_cnt),0) AS sr_acad_mnths, nullif(SUM(sr_cal_mnth_cnt),0) AS sr_cal_mnths
INTO #myPropBudg FROM #myProp p
JOIN #myBudg eb ON p.prop_id = eb.prop_id 
GROUP BY p.prop_id, eb.revn_num
ORDER BY p.prop_id -- eb.revn_num eb.budg_seq_yr
CREATE INDEX myPropBudg_ix ON #myPropBudg(prop_id)
-- per-proposal data 
SELECT pid.*, rqst_dol as D, prc.R, b.T, b.RN, id=identity(18), 0 as seq 
INTO #myPropInfo FROM #myProp pid
JOIN csd.prop p ON p.prop_id = pid.prop_id
LEFT JOIN #myPRCs prc ON prc.prop_id = pid.prop_id
LEFT JOIN #myPropBudg b ON b.prop_id = p.prop_id
ORDER BY lead, ILN, pid.prop_id
CREATE INDEX myProp_idx ON #myPropInfo(prop_id)
SELECT lead, MIN(id) as start INTO #myPropStart FROM #myPropInfo GROUP BY lead
UPDATE #myPropInfo SET seq = id-M.start FROM #myPropInfo r, #myPropStart M WHERE r.lead = M.lead
DROP TABLE #myPropStart
CREATE INDEX myProp_ix ON #myPropInfo(lead)
--]RA_prop
--select * from #myPropBudg select * from #myPropInfo

--Review scores: rev_prop and rev_prop_vw are the eJ & Fastlane database tables
-- rp holds status  rpv holds split scores and release flags
-- pds can update rp score (but can't split), so disagreements are hard to adjudicate.
-- Deletion from FL propagates overnight to eJ. (FL subm flag = D not to be confused with eJ stts = D)
--Assumptions: eJ has the right review score except when it is R, when the FL score is split, or when the FL subm is D. (Check rev score flag.)
--  rev_prop stts C should never be shown. Include in text helper sans score. 

-- first, get all form 7 prop, panl, revr assignments that matter (C or scored.)
-- For reviews, take FL score/string if it exists, else eJ score/string; also takes all conflicted and released reviews/assignments.
--  the hazzard is that FL reviews can be corrected in eJ, but since this requires PO action, the PO can write it in the RA.
--  (The problem with the other way is that eJ can't record split scores, and in some instances (GEO EAR) records assignments that have no scores, 
--   so it is really hard to determine which deserves to count.  Choosing FL first makes an easy to state policy.)
-- panelists on >1 panels; review credited to first panel.

-- rp.rev_prop_rtng_code, rp.rev_stts_code, 
--rpv.rev_rlse_flag, rpv.rev_prop_unrl_flag, rpv.rev_subm_flag,
-- CASE WHEN rpv.rev_prop_rtng_ind IS NOT NULL THEN 1 ELSE 0 END as rcvdFL,
 
-- 0/1 parameters allow summing later. We want to know:
--  When are we releasing Conflicted, Pending, Selected, etc reviews
--  How many are marked unreleasable, have unmarked FL reviews, are pending or selected
--  When are the eJ ratings different from the FL ratings


--select rp.*, rpv.* from csd.rev_prop rp
--LEFT JOIN csd.rev_prop_vw rpv ON rpv.revr_id = rp.revr_id AND rpv.prop_id = rp.prop_id  
--where rp.prop_id = '1651952'
--select * from #myProp where prop_id = '1749977'
--select * from #myRevs where lead = '1651952'
--select * from csd.prop where prop_id = '1749977'

--select * from #myRevs where diffFLeJ = 1
--select * from #myRevs where pendSlct = 1
--select * from #myRevs where rlsdCDNPS = 1
--select * from #myRevs where unmkd = 1

--[RA_revs
-- Needs: RA_lead(#myLead),
--DROP TABLE #myRevs, #myRevMarks, #myRevSumm, #myRevPanl
declare @adhoc char(7), @olddate datetime
SELECT @adhoc = '.ad hoc ',  @olddate = '1/1/2000' -- for formatting dates
SELECT isnull((SELECT MIN(pp.panl_id) FROM csd.panl_prop pp 
    JOIN csd.panl_revr pr ON pr.panl_id = pp.panl_id AND pp.panl_id like 'P%'
    WHERE pp.prop_id = p.lead AND pr.revr_id = rp.revr_id),  @adhoc) as panl_id, p.lead, rp.revr_id, rp.rev_rtrn_date, 
 CASE WHEN rs.score IS NOT NULL THEN rs.string ELSE rp.rev_prop_rtng_code END as string,
 CASE WHEN rs.score IS NOT NULL THEN rs.score ELSE CASE rp.rev_prop_rtng_code WHEN 'E' THEN 9 WHEN 'V' THEN 7 
    WHEN 'G' THEN 5 WHEN 'F' THEN 3 WHEN 'P' THEN 1 ELSE NULL END END AS score,
 CASE WHEN rp.rev_stts_code = 'C' THEN 1 ELSE 0 END as confl,
 CASE WHEN rp.rev_stts_code IN ('P','S') THEN 1 ELSE 0 END as pendSlct,
 CASE WHEN rpv.rev_prop_unrl_flag = 'Y' THEN 1 ELSE 0 END as unrlsbl,
 CASE WHEN rpv.rev_prop_rtng_ind IS NULL OR rpv.rev_rlse_flag = 'Y' OR rpv.rev_prop_unrl_flag = 'Y' THEN 0 ELSE 1 END as unmkd,
 CASE WHEN rpv.rev_rlse_flag = 'Y' AND rp.rev_stts_code IN ('C','D','N','P','S') THEN 1 ELSE 0 END as rlsdCDNPS,
 CASE WHEN rp.rev_stts_code IN ('C','D','N','R') OR rp.rev_prop_rtng_code = nullif(rs.string,' ') THEN 0 ELSE 1 END as diffFLeJ,
 id=identity(18), 0 as seq
INTO #myRevs FROM #myLead p
JOIN csd.rev_prop rp ON rp.prop_id = p.lead
LEFT JOIN csd.rev_prop_vw rpv ON rpv.revr_id = rp.revr_id AND rpv.prop_id = rp.prop_id  
LEFT JOIN tempdb.guest.revScores rs ON rs.yn = rpv.rev_prop_rtng_ind -- this table uses the same 1-9 scale above, but handles split scores
WHERE (rpv.rev_rlse_flag = 'Y' OR rp.rev_stts_code = 'C' OR rp.rev_prop_rtng_code IN ('E','V','G','F','P') -- has eJ review
       OR rpv.rev_prop_rtng_ind > 'NNNNN') -- has FL review.  checked: no stts N or D come in.  R,S,P do, all good.
  AND isnull(rpv.rev_subm_flag,'U') <> 'D' -- ignore reviews deleted on FL, even if that takes overnight to propagate to eJ.
ORDER BY lead, confl, score DESC, revr_id -- move C last
CREATE INDEX myRevs_ix0 ON #myRevs(lead)
CREATE INDEX myRevs_ix1 ON #myRevs(lead, confl)
CREATE INDEX myRevs_ix2 ON #myRevs(panl_id, lead, confl)
SELECT lead, MIN(id) as 'start' INTO #myStarts FROM #myRevs GROUP BY lead
UPDATE #myRevs SET seq = id-M.start FROM #myRevs r, #myStarts M WHERE r.lead = M.lead 
DROP TABLE #myStarts

SELECT lead,  nullif(sum(confl),0) as Nconfl,
nullif(sum(unrlsbl),0) as Nunrlsbl,nullif(sum(rlsdCDNPS),0) as NrlsdCDNPS, 
nullif(sum(pendSlct),0) as NpendSlct, nullif(sum(diffFLeJ),0) as NdiffFLeJ
INTO #myRevMarks FROM #myRevs
GROUP BY lead
CREATE INDEX myRevMarks_ix ON #myRevMarks(lead)

SELECT lead, count(DISTINCT revr_id) as Nrev, nullif(sum(unmkd),0) as Nunmkd,
MIN(r.score) AS minScore, AVG(r.score) AS avg_score,MAX(r.score) AS maxScore,
convert(varchar(50), MAX(CASE r.seq WHEN  0 THEN     r.string ELSE '' END)+
 MAX(CASE r.seq WHEN  1 THEN ','+r.string ELSE '' END)+
 MAX(CASE r.seq WHEN  2 THEN ','+r.string ELSE '' END)+
 MAX(CASE r.seq WHEN  3 THEN ','+r.string ELSE '' END)+
 MAX(CASE r.seq WHEN  4 THEN ','+r.string ELSE '' END)+
 MAX(CASE r.seq WHEN  5 THEN ','+r.string ELSE '' END)+
 MAX(CASE r.seq WHEN  6 THEN ','+r.string ELSE '' END)+
 MAX(CASE r.seq WHEN  7 THEN ','+r.string ELSE '' END)+
 MAX(CASE r.seq WHEN  8 THEN ','+r.string ELSE '' END)+
 MAX(CASE r.seq WHEN  9 THEN ','+r.string ELSE '' END)+
 MAX(CASE r.seq WHEN 10 THEN ','+r.string ELSE '' END)+
 MAX(CASE r.seq WHEN 11 THEN ','+r.string ELSE '' END)+
 MAX(CASE r.seq WHEN 12 THEN ','+r.string ELSE '' END)+
 MAX(CASE r.seq WHEN 13 THEN ','+r.string ELSE '' END)+
 MAX(CASE r.seq WHEN 14 THEN ','+r.string ELSE '' END)) AS allReviews, 
MAX(r.rev_rtrn_date) AS last_rev_date
INTO #myRevSumm FROM #myRevs r WHERE confl = 0
GROUP BY r.lead
CREATE INDEX myRevSumm_ix ON #myRevSumm(lead)

SELECT lead, panl_id, count(revr_id) as N, convert(varchar(50), STUFF(LTRIM(
 MAX(CASE r.seq WHEN  0 THEN ','+r.string ELSE '' END)+
 MAX(CASE r.seq WHEN  1 THEN ','+r.string ELSE '' END)+
 MAX(CASE r.seq WHEN  2 THEN ','+r.string ELSE '' END)+
 MAX(CASE r.seq WHEN  3 THEN ','+r.string ELSE '' END)+
 MAX(CASE r.seq WHEN  4 THEN ','+r.string ELSE '' END)+
 MAX(CASE r.seq WHEN  5 THEN ','+r.string ELSE '' END)+
 MAX(CASE r.seq WHEN  6 THEN ','+r.string ELSE '' END)+
 MAX(CASE r.seq WHEN  7 THEN ','+r.string ELSE '' END)+
 MAX(CASE r.seq WHEN  8 THEN ','+r.string ELSE '' END)+
 MAX(CASE r.seq WHEN  9 THEN ','+r.string ELSE '' END)+
 MAX(CASE r.seq WHEN 10 THEN ','+r.string ELSE '' END)+
 MAX(CASE r.seq WHEN 11 THEN ','+r.string ELSE '' END)+
 MAX(CASE r.seq WHEN 12 THEN ','+r.string ELSE '' END)+
 MAX(CASE r.seq WHEN 13 THEN ','+r.string ELSE '' END)+
 MAX(CASE r.seq WHEN 14 THEN ','+r.string ELSE '' END)),1,1,'')) AS V, 
MAX(r.rev_rtrn_date) AS last_rev_date
INTO #myRevPanl FROM #myRevs r WHERE confl = 0
GROUP BY lead, panl_id
CREATE INDEX myRevPanl_ix ON #myRevPanl(lead, panl_id)
--]RA_revs

--per project conflicts recorded
--[RA_confl
--Needs: #myRevs, #myRevPanl
--DROP TABLE #myPPConfl
select rp.lead, rp.panl_id, rtrim(r.revr_frst_name) + ' ' + rtrim(r.revr_last_name) +', ' + rtrim(i.inst_shrt_name) as confl,
--c.revr_id, r.revr_last_name, r.revr_frst_name, i.inst_shrt_name, 
id=identity(18), 0 as seq
INTO #myPConfl FROM #myRevPanl rp
JOIN #myRevs c ON c.panl_id = rp.panl_id and c.lead = rp.lead and c.confl=1
JOIN csd.revr r ON r.revr_id = c.revr_id
LEFT JOIN csd.inst i ON i.inst_id = r.inst_id
ORDER BY rp.lead, rp.panl_id, r.revr_last_name

SELECT lead, panl_id, MIN(id) as start INTO #myCStarts FROM #myPConfl GROUP BY lead, panl_id
UPDATE #myPConfl SET seq = id-M.start FROM #myPConfl r, #myCStarts M  WHERE r.lead = M.lead and r.panl_id = M.panl_id
DROP TABLE #myCStarts

SELECT lead, panl_id, count(confl) as C,  
 MAX(CASE r.seq WHEN  0 THEN r.confl ELSE '' END)+
 MAX(CASE r.seq WHEN  1 THEN '; '+r.confl ELSE '' END)+
 MAX(CASE r.seq WHEN  2 THEN ','+r.confl ELSE '' END)+
 MAX(CASE r.seq WHEN  3 THEN ','+r.confl ELSE '' END)+
 MAX(CASE r.seq WHEN  4 THEN ','+r.confl ELSE '' END)+
 MAX(CASE r.seq WHEN  5 THEN ','+r.confl ELSE '' END)+
 MAX(CASE r.seq WHEN  6 THEN ','+r.confl ELSE '' END) AS CR
INTO #myPPConfl  FROM #myPConfl r 
GROUP BY lead, panl_id
CREATE INDEX myPPConfl_ix ON #myPPConfl(lead, panl_id)
DROP TABLE #myPConfl
--]RA_confl
--select * from #myPConfl


-- panel outcomes
-- The information on how other projects fared in a panel is hard to get except by sql query, so this is valuable info for RAs.
-- The #myPanlOutcomes queries identify all relevant panels, then pull recommendations for ALL proposals on the panels and concatenates.
-- Conversions like null to varchar(126) make sure that the types are long enough to take data that may come later.
-- Ranks come from Interactive Panel System (IPS); since some PDs don't enter ranks there, need to fail gracefully if none are supplied.
--[RA_panl
-- Needs: RA_lead(#myLead),RA_revs(#myRevs, #myRevPanl)
--DROP TABLE #myPanl, #myProjPanl, #myProjPanlSumm
SELECT pn.panl_id AS I, panl_name AS PN, panl_end_date AS E, 
  (SELECT COUNT(DISTINCT revr_id) FROM csd.panl_revr r WHERE r.panl_id = pn.panl_id) AS P,
  convert(varchar(126),NULL) AS S 
INTO #myPanl FROM (SELECT DISTINCT panl_id FROM #myLead pid JOIN csd.panl_prop pp ON pp.prop_id = pid.lead) pn
JOIN csd.panl panl ON panl.panl_id = pn.panl_id
CREATE INDEX myPanl_ix ON #myPanl(I)
SELECT pp.panl_id, s.RCOM_SEQ_NUM, RCOM_ABBR, Count(s.PROP_ID) AS rtCount
INTO #myPanlOutcomes FROM #myPanl pl
JOIN csd.panl_prop pp ON pp.panl_id = pl.I -- for all lead props on my panls
JOIN csd.prop p ON p.prop_id = pp.prop_id AND isnull(lead_prop_id,p.prop_id) = p.prop_id
JOIN FLflpdb.flp.panl_prop_summ s ON s.PANL_ID = pl.I AND s.PROP_ID = pp.prop_id
LEFT JOIN FLflpdb.flp.panl_rcom_def d ON d.PANL_ID = pl.I AND d.RCOM_SEQ_NUM = s.RCOM_SEQ_NUM
GROUP BY pp.panl_id, s.RCOM_SEQ_NUM, RCOM_ABBR
ORDER BY pp.panl_id, s.RCOM_SEQ_NUM
UPDATE #myPanl SET S = (SELECT isnull(convert(varchar,SUM(ps.rtCount)),'No') + ' rated projects' +
    MAX( CASE ps.RCOM_SEQ_NUM WHEN 1 THEN ': ' + convert(varchar,ps.rtCount) + ' ' +  ps.RCOM_ABBR ELSE '' END ) +
    MAX( CASE ps.RCOM_SEQ_NUM WHEN 2 THEN ', ' + convert(varchar,ps.rtCount) + ' ' +  ps.RCOM_ABBR ELSE '' END ) +
    MAX( CASE ps.RCOM_SEQ_NUM WHEN 3 THEN ', ' + convert(varchar,ps.rtCount) + ' ' +  ps.RCOM_ABBR ELSE '' END ) +
    MAX( CASE ps.RCOM_SEQ_NUM WHEN 4 THEN ', ' + convert(varchar,ps.rtCount) + ' ' +  ps.RCOM_ABBR ELSE '' END ) +
    MAX( CASE ps.RCOM_SEQ_NUM WHEN 5 THEN ', ' + convert(varchar,ps.rtCount) + ' ' +  ps.RCOM_ABBR ELSE '' END ) +
    MAX( CASE ps.RCOM_SEQ_NUM WHEN 6 THEN ', ' + convert(varchar,ps.rtCount) + ' ' +  ps.RCOM_ABBR ELSE '' END )
    FROM #myPanlOutcomes ps WHERE ps.panl_id = pl.I) 
FROM #myPanl pl DROP TABLE #myPanlOutcomes

--per project panel summary 
SELECT rp.lead, ps.*, rp.N, rp.V, s.RCOM_SEQ_NUM AS RS, d.RCOM_ABBR as RA, d.RCOM_TXT as RT, s.PROP_ORDR as RK,
c.C, C.CR,
--nullif((SELECT count(*) FROM #myRevs c WHERE confl=1 AND c.panl_id = rp.panl_id AND c.lead = rp.lead ),0) as C2,
CASE WHEN panl_summ_unrl_flag = 'Y' THEN 1 ELSE 0 END as summ_unrls, 
CASE WHEN panl_summ_unrl_flag = 'Y' OR panl_summ_rlse_flag = 'Y'  THEN 0 ELSE 1 END as summ_unmrkd, id=identity(18), 0 as seq
INTO #myProjPanl FROM #myRevPanl rp
JOIN #myPanl ps ON ps.I = rp.panl_id
LEFT JOIN #myPPConfl c ON c.lead = rp.lead AND c.panl_id = rp.panl_id
LEFT JOIN FLflpdb.flp.panl_prop_summ s ON s.panl_id = rp.panl_id AND s.prop_id = rp.lead
LEFT JOIN FLflpdb.flp.panl_rcom_def d ON d.panl_id = rp.panl_id  AND d.RCOM_SEQ_NUM = s.RCOM_SEQ_NUM
ORDER BY lead, ps.E
CREATE INDEX myProjPanl_ix ON #myProjPanl(lead)
SELECT lead, MIN(id) as start INTO #myPStarts FROM #myProjPanl GROUP BY lead
UPDATE #myProjPanl SET seq = id-M.start FROM #myProjPanl r, #myPStarts M  WHERE r.lead = M.lead
DROP TABLE #myPStarts
SELECT lead, count(I) AS nPanl, min(RS)+isnull(min(RK),0)/100.0 AS RecRkMin, nullif(sum(summ_unrls),0) as nPSunrls, 
    nullif(sum(summ_unmrkd),0) as nPSunmrkd 
INTO #myProjPanlSumm FROM #myProjPanl GROUP BY lead
--]RA_panl
--select * from #myPanl select * from #myProjPanl select * from #myProjPanlSumm 
--select lead,C,C2,CR from #myProjPanl 
--
---- main query steps -- mostly deleted, but available on github
--SELECT p.lead,
--ra.RAupdate
--FROM #myLead p
--LEFT JOIN #myRA ra ON ra.lead = p.lead

--declare @adhoc char(7), @olddate datetime  SELECT @adhoc = '.ad hoc ',  @olddate = '1/1/2000' -- for formatting dates
--[RA_allRAdata
-- Needs: RA_lead(#myLead,#myRA),RA_propPRCs(#myPRCs),RA_prop(#myPropInfo,#myPropBudg),RA_revs(#myRevs, #myRevMarks, #myRevPanl), RA_panl(#myPanl, #myProjPanl, #myProjPanlSumm)
SELECT nullif(dd_rcom_date,'1900-01-01') AS dd_rcom_date, nsf_rcvd_date, 
prop.pgm_annc_id, prop.org_code, prop.pgm_ele_code, natr_rqst.natr_rqst_abbr, prop_stts_abbr, prop.obj_clas_code,
cntx_stmt_id,p.PO,ra.RAupdate, 
nPanl, isnull(RecRkMin,99)+nullif(avg_score,0)/1000 AS RecRkMin, pn0.RA as rec0, pn1.RA as rec1, pn2.RA as rec2, 
pid.RAtemplate, org.dir_div_abbr as Div, p.lead, Nrev,Nunmkd,minScore,avg_score,maxScore,allReviews,rs.last_rev_date, 
rm.Nunrlsbl, projTot.rqst_tot, budg_tot, budRevnMax, 
prop.rqst_eff_date, prop.rqst_mnth_cnt, 
rtrim(prop_titl_txt) AS prop_titl_txt, 
ah.N as AhNrev, ah.V as AhRevs, ah.last_rev_date as AhLast,
 pn0.I AS panl_id0, pn0.RT AS RCOM_TXT0, pn0.RK AS rank0, pn0.PN AS panl_name0,pn0.E AS panl_end0,pn0.V AS revs0,pn0.S AS PanlString0,pn0.P as pnlst0, pn0.C as confl0, pn0.CR as conflrev0,
 pn1.I AS panl_id1, pn1.RT AS RCOM_TXT1, pn1.RK AS rank1, pn1.PN AS panl_name1,pn1.E AS panl_end1,pn1.V AS revs1,pn1.S AS PanlString1,pn1.P as pnlst1, pn1.C as confl1, pn1.CR as conflrev1,
 pn2.I AS panl_id2, pn2.RT AS RCOM_TXT2, pn2.RK AS rank2, pn2.PN AS panl_name2,pn2.E AS panl_end2,pn2.V AS revs2,pn2.S AS PanlString2,pn2.P as pnlst2, pn2.C as confl2, pn2.CR as conflrev2,
p0.prop_id as prop_id0, p0.L as last0, p0.F as frst0, p0.I as inst0, p0.D as rqst0, p0.T as b0tot, p0.R as PRC0,
p1.prop_id as prop_id1, p1.L as last1, p1.F as frst1, p1.I as inst1, p1.D as rqst1, p1.T as b1tot, p1.R as PRC1,
p2.prop_id as prop_id2, p2.L as last2, p2.F as frst2, p2.I as inst2, p2.D as rqst2, p2.T as b2tot, p2.R as PRC2,
p3.prop_id as prop_id3, p3.L as last3, p3.F as frst3, p3.I as inst3, p3.D as rqst3, p3.T as b3tot, p3.R as PRC3,
p4.prop_id as prop_id4, p4.L as last4, p4.F as frst4, p4.I as inst4, p4.D as rqst4, p4.T as b4tot, p4.R as PRC4,
p5.prop_id as prop_id5, p5.L as last5, p5.F as frst5, p5.I as inst5, p5.D as rqst5, p5.T as b5tot, p5.R as PRC5,
p6.prop_id as prop_id6, p6.L as last6, p6.F as frst6, p6.I as inst6, p6.D as rqst6, p6.T as b6tot, p6.R as PRC6,
rtrim(pa.dflt_prop_titl_txt) AS solicitation, org.org_long_name as Div_name, 
pgm_ele_name, sign_blck_name, prop_stts_txt, natr_rqst_txt, obj_clas_name,
o2.dir_div_abbr as Dir, o2.org_long_name as Dir_name, getdate() AS pulldate, 
p0.M AS email, convert(varchar(255),(SELECT MAX(CASE r.seq WHEN  0 THEN r.M ELSE '' END)+
 MAX(CASE r.seq WHEN  1 THEN ';'+r.M ELSE '' END)+
 MAX(CASE r.seq WHEN  2 THEN ';'+r.M ELSE '' END)+
 MAX(CASE r.seq WHEN  3 THEN ';'+r.M ELSE '' END)+
 MAX(CASE r.seq WHEN  4 THEN ';'+r.M ELSE '' END)+
 MAX(CASE r.seq WHEN  5 THEN ';'+r.M ELSE '' END)+
 MAX(CASE r.seq WHEN  6 THEN ';'+r.M ELSE '' END) 
FROM #myPropInfo r WHERE r.lead = p.lead)) AS allPIemail
--INTO #myTmp -- save to get types for formatting dummy line
FROM #myLead p
JOIN #myPidRAt pid ON pid.prop_id = p.lead
JOIN csd.prop prop ON prop.prop_id = p.lead
JOIN FLflpdb.flp.org org ON org.org_code = prop.org_code
LEFT JOIN FLflpdb.flp.org o2 ON o2.org_code =left(prop.org_code,2)+'000000' 
LEFT JOIN FLflpdb.flp.pgm_ele pe ON pe.pgm_ele_code = prop.pgm_ele_code
LEFT JOIN csd.pgm_annc pa ON pa.pgm_annc_id = prop.pgm_annc_id
JOIN FLflpdb.flp.obj_clas_pars oc ON oc.obj_clas_code = prop.obj_clas_code
JOIN csd.prop_stts prop_stts ON prop_stts.prop_stts_code = prop.prop_stts_code
JOIN csd.natr_rqst natr_rqst ON natr_rqst.natr_rqst_code = prop.natr_rqst_code
LEFT JOIN #myRA ra ON ra.lead = p.lead
LEFT JOIN #myRevSumm rs ON rs.lead = p.lead
LEFT JOIN #myRevMarks rm ON rm.lead = p.lead
LEFT JOIN csd.po_vw po_vw ON po_vw.po_ibm_logn_id = p.PO
LEFT JOIN #myRevPanl ah ON ah.lead = p.lead AND ah.panl_id =  @adhoc
LEFT JOIN (SELECT * FROM #myProjPanl WHERE 0=seq) pn0 ON pn0.lead = p.lead
LEFT JOIN (SELECT * FROM #myProjPanl WHERE 1=seq) pn1 ON pn1.lead = p.lead
LEFT JOIN (SELECT * FROM #myProjPanl WHERE 2=seq) pn2 ON pn2.lead = p.lead 
LEFT JOIN (SELECT lead, count(I) AS nPanl, min(RS)+isnull(min(RK),0)/100.0 AS RecRkMin 
           FROM #myProjPanl GROUP BY lead) pn ON pn.lead = p.lead
LEFT JOIN (SELECT * FROM #myPropInfo mp WHERE mp.seq = 0) p0 ON p0.lead = p.lead
LEFT JOIN (SELECT * FROM #myPropInfo mp WHERE mp.seq = 1) p1 ON p1.lead = p.lead
LEFT JOIN (SELECT * FROM #myPropInfo mp WHERE mp.seq = 2) p2 ON p2.lead = p.lead
LEFT JOIN (SELECT * FROM #myPropInfo mp WHERE mp.seq = 3) p3 ON p3.lead = p.lead
LEFT JOIN (SELECT * FROM #myPropInfo mp WHERE mp.seq = 4) p4 ON p4.lead = p.lead
LEFT JOIN (SELECT * FROM #myPropInfo mp WHERE mp.seq = 5) p5 ON p5.lead = p.lead
LEFT JOIN (SELECT * FROM #myPropInfo mp WHERE mp.seq = 6) p6 ON p6.lead = p.lead
LEFT JOIN (SELECT lead, SUM(D) AS rqst_tot, SUM(T) AS budg_tot, MAX(RN) AS budRevnMax
           FROM #myPropInfo GROUP BY lead) projTot ON projTot.lead = p.lead
--]RA_allRAdata
--[RA_allRAdata2
UNION ALL SELECT @olddate,@olddate -- example to set mail merge format.  
,'.first line.','..please','keep','for ','mail','merge','formatting.','thx.',@olddate
,0,-0.99999999,'NRFP','NRFP','NRFP','zztemplateFname RAt.docx'
,'Divn','1234567',0,0,1,4.5,9,'E,E/V,V,V/G,G,G/F,F,F/P,P',@olddate, 0
,99999999.0,99999999.0,0,@olddate,36
,'This first row is needed so that mail merge formatting will be correct. Please do not remove it.  Mail merge takes its formatting from the first rows of the table............'
,0,'E,V,G,F,P',@olddate
,'P123456','Dont remove this line for panel rec formatting....',99,'Panel name spelled out; keep this line for formatting.......',@olddate,'E,E/V,V,V/G,G,G/F,F,F/P,P','Reports rank & all recs (HC, etc.) from Interactive Panel System (IPS).  I suggest Standard competition rank  (wikipedia) .',0,0,'List of conflicted panelists separated by semicolons. Not everyone will need this list, but some divisions request this information, so I am seeing if we can provide it. There is a limit of seven names returned, though all will be counted.'
,'P123456','Dont remove this line for panel rec formatting....',99,'Panel name spelled out; keep this line for formatting.......',@olddate,'E,E/V,V,V/G,G,G/F,F,F/P,P','Reports rank & all recs (HC, etc.) from Interactive Panel System (IPS).  I suggest Standard competition rank  (wikipedia) .',0,0,'List of conflicted panelists separated by semicolons. Not everyone will need this list, but some divisions request this information, so I am seeing if we can provide it. There is a limit of seven names returned, though all will be counted.'
,'P123456','Dont remove this line for panel rec formatting....',99,'Panel name spelled out; keep this line for formatting.......',@olddate,'E,E/V,V,V/G,G,G/F,F,F/P,P','Reports rank & all recs (HC, etc.) from Interactive Panel System (IPS).  I suggest Standard competition rank  (wikipedia) .',0,0,'List of conflicted panelists separated by semicolons. Not everyone will need this list, but some divisions request this information, so I am seeing if we can provide it. There is a limit of seven names returned, though all will be counted.'
,'1234567','PI last name for format.','PI first name..','Inst name for formatting',99999999.0,99999999.0,'Proposal PRCs assgnd; see glossary '
,'1234567','PI last name for format.','PI first name..','Inst name for formatting',99999999.0,99999999.0,'Proposal PRCs assgnd; see glossary '
,'1234567','PI last name for format.','PI first name..','Inst name for formatting',99999999.0,99999999.0,'Proposal PRCs assgnd; see glossary '
,'1234567','PI last name for format.','PI first name..','Inst name for formatting',99999999.0,99999999.0,'Proposal PRCs assgnd; see glossary '
,'1234567','PI last name for format.','PI first name..','Inst name for formatting',99999999.0,99999999.0,'Proposal PRCs assgnd; see glossary '
,'1234567','PI last name for format.','PI first name..','Inst name for formatting',99999999.0,99999999.0,'Proposal PRCs assgnd; see glossary '
,'1234567','PI last name for format.','PI first name..','Inst name for formatting',99999999.0,99999999.0,'Proposal PRCs assgnd; see glossary '
,'Solicitation name retrieved from pgm_annc. This example is for formatting; please do not remove this line............'
,'Division or directorate name retrieved by org_code from org. This example is for formatting; please do not remove this line............'
,'Program Element name retrieved','Name for PO signature','Proposal status details','Nature of request full name','Object Class full name','DIR'
,'Directorate name retrieved by modified org_code from org. This example is for formatting; please do not remove this line............'
,@olddate,'Email of lead PI on the project','list of all emails for Pis on Lead and non-lead proposals on the project.  Does not include the co-Pis. '
--]RA_allRAdata2

-- To get types for formatting, uncomment this line above --INTO #myTmp 
--    & run query (ignore warning about row size)
-- Then in tempdb run this query:
--    SELECT sc.colid, sc.name, t.name as type, sc.length FROM sysobjects so 
--    JOIN syscolumns sc ON sc.id = so.id 
--    JOIN systypes t on t.usertype = sc.usertype
--    WHERE so.name like '#myTmp%' -- table name
--    ORDER BY colid
-- Check the types and lengths are expected, then put the results in a table in Excel and format.
-- e.g,quote char & enforce max length: 
--    =IF(RIGHT([[type]],4)="char",",'"&LEFT([[content]],[[length]])&"'",","&[[content]])
-- Comment out --INTO again. DROP TABLE #myTmp

-------------
-- PRC Glossary
-- We'll do this query first to fill the caches, 
-- so this is a good place to put the check for revtable.
--[revtable
--NB: Should complete before queries that need table revScores to handle split review scores in subsequent queries.
if object_id('tempdb.guest.revScores') is null 
  exec('create table tempdb.guest.revScores(yn char(5) primary key, string varchar(10), score real null) insert into tempdb.guest.revScores
select ''NNNNN'', ''R'',  null  union all select ''NNNNY'', ''P'', 1 union all 
select ''NNNYN'', ''F'',    3   union all select ''NNNYY'', ''F/P'', 2 union all
select ''NNYNN'', ''G'',    5   union all select ''NNYNY'', ''G/P'', 2.98 union all
select ''NNYYN'', ''G/F'',  4   union all select ''NNYYY'', ''G/F/P'', 2.99 union all
select ''NYNNN'', ''V'',    7   union all select ''NYNNY'', ''V/P'', 3.98 union all
select ''NYNYN'', ''V/F'', 4.98 union all select ''NYNYY'', ''V/F/P'', 3.65 union all
select ''NYYNN'', ''V/G'',  6   union all select ''NYYNY'', ''V/G/P'', 4.32 union all
select ''NYYYN'',''V/G/F'',4.99 union all select ''NYYYY'', ''V/G/F/P'', 3.97 union all
select ''YNNNN'', ''E'',    9   union all select ''YNNNY'', ''E/P'', 4.992 union all
select ''YNNYN'', ''E/F'', 5.98 union all select ''YNNYY'', ''E/F/P'', 4.325 union all
select ''YNYNN'', ''E/G'', 6.98 union all select ''YNYNY'', ''E/G/P'', 4.995 union all
select ''YNYYN'',''E/G/F'',5.66 union all select ''YNYYY'', ''E/G/F/P'', 4.5 union all
select ''YYNNN'', ''E/V'',  8   union all select ''YYNNY'', ''E/V/P'', 5.666 union all
select ''YYNYN'',''E/V/F'',6.33 union all select ''YYNYY'', ''E/V/F/P'', 4.996 union all
select ''YYYNN'',''E/V/G'',6.99 union all select ''YYYNY'', ''E/V/G/P'', 5.5 union all
select ''YYYYN'', ''E/V/G/F'', 5.99 union all select ''YYYYY'', ''E/V/G/F/P'', 4.997')
--]revtable

--[RA_PRCglossary
--NON-PUBLIC TABLE: csd.pgm_ref
--Needs: RA_lead, RA_propPRCs(#myPRCdata)
SELECT p.prop_atr_code AS PRC, prc.pgm_ref_txt AS PRC_Description
FROM (SELECT DISTINCT prop_atr_code FROM #myPRCdata) p
JOIN csd.pgm_ref prc ON prc.pgm_ref_code = p.prop_atr_code
ORDER BY PRC
--]RA_PRCglossary
-----------------
--[RA_no_PRC
--Needs: RA_lead, RA_propPRCs(#myPRCdata)
SELECT p.prop_atr_code AS PRC, 'table rptdb.csd.pgm_ref missing at install' AS PRC_Description
FROM (SELECT DISTINCT prop_atr_code FROM #myPRCdata) p
ORDER BY PRC
--]RA_no_PRC
-----------------

--Checkers

-- check coding
-- Things to do in excel
--check epscor:
-- =XOR(IFERROR(MATCH([@[st_code]],EpscorTable,0)>0,FALSE),IFERROR(FIND("9150",[@propPRCs])>0,FALSE))
--check WMD
-- =IF(SUM(--dmog columns--)>0,"WMD","")
-- =XOR([@WMD]<>"",IFERROR(FIND("9102",[@propPRCs])>0,FALSE))
-- number by proj:
-- =S1+([ILN]<"M")
-- eJ hyperlink:
-- =HYPERLINK("https://www.ejacket.nsf.gov/ej/showProposal.do?ID="&[@[prop_id]],"eJ")
-- mailto link:
-- = HYPERLINK("mailto:"&[@email]&"&subject=Your NSF Proposal:  "&[@[prop_id]],[@email])
-- coloring rsch,other % don't add to 1
-- nonleads differ from line above
-- titles differ -- except whitespace
-- summaries, reviews released
-- conflicted reviews

--[RA_propCheck
-- Needs: RA_lead(#myProp,#myRA),RA_prop(#myPropInfo,#myPropBudg),RA_revs(#myRevMarks, #myRevPanlSumm), RA_panl(#myProjPanlSumm)
-- DROP TABLE #myDmog
SELECT prop_id, 
SUM(CASE WHEN pi_gend_code = 'F'THEN 1 ELSE 0 END) AS NfmlPIs,
SUM(CASE WHEN pi_ethn_code = 'H'THEN 1 ELSE 0 END) AS NhispPIs,
SUM(CASE WHEN dmog_tbl_code = 'H'AND dmog_code <> 'N' THEN 1 ELSE 0 END) AS NhndcpPIs,
SUM(CASE WHEN dmog_tbl_code = 'R'AND dmog_code NOT IN ('U','W','B3') THEN 1 ELSE 0 END) AS NnonWhtAsnPIs
INTO #myDmog FROM (SELECT p.prop_id, pi_id FROM #myProp p 
      UNION ALL SELECT p.prop_id, a.pi_id FROM #myProp p JOIN csd.addl_pi_invl a ON a.prop_id = p.prop_id) PIs
LEFT JOIN csd.pi_vw pi ON pi.pi_id = PIs.pi_id
LEFT JOIN csd.PI_dmog d ON d.pi_id = PIs.pi_id
GROUP BY prop_id
ORDER BY prop_id
CREATE INDEX myDmog_idx ON #myDmog(prop_id)

-- props: get codes to check if they match leads
SELECT nullif(dd_rcom_date,'1900-01-01') AS dd_rcom_date, nsf_rcvd_date,
prop.pgm_annc_id, prop.org_code, prop.pgm_ele_code,
prop_stts_abbr,natr_rqst.natr_rqst_abbr,prop.obj_clas_code,
bas_rsch_pct, apld_rsch_pct+educ_trng_pct+land_buld_fix_equp_pct+mjr_equp_pct+non_invt_pct AS other_pct,  
cntx_stmt_id, p.PO, p.R as propPRCs, p.st_code, ra.RAupdate,
p.lead, p.ILN, --eJ link
org.dir_div_abbr as Div, p.prop_id,
p.L AS pi_last_name, p.F AS pi_frst_name, p.I AS inst, --mailto:
prop.rqst_eff_date,prop.rqst_mnth_cnt,
nPanl, nPSunmrkd, nPSunrls, RecRkMin, 
rs.nRev, rs.allReviews, rs.avg_score, rs.last_rev_date,
Nunmkd, Nconfl, Nunrlsbl, NrlsdCDNPS, NpendSlct, NdiffFLeJ,
 prop_stts_txt, prop.prop_titl_txt,
p.D AS rqst_tot, NfmlPIs,NhispPIs,NhndcpPIs,NnonWhtAsnPIs,
p.M as email
FROM #myPropInfo p
JOIN csd.prop prop ON prop.prop_id = p.prop_id
JOIN FLflpdb.flp.org org ON org.org_code = prop.org_code
JOIN csd.prop_stts prop_stts ON prop_stts.prop_stts_code = prop.prop_stts_code
JOIN csd.natr_rqst natr_rqst ON natr_rqst.natr_rqst_code = prop.natr_rqst_code
LEFT JOIN #myRA ra ON ra.lead = p.lead
LEFT JOIN #myDmog d ON d.prop_id = p.prop_id
LEFT JOIN #myRevSumm rs ON rs.lead = p.prop_id
LEFT JOIN #myRevMarks rv ON rv.lead = p.prop_id
LEFT JOIN #myProjPanlSumm pl ON pl.lead = p.prop_id
ORDER BY p.lead, p.ILN, p.prop_id
--]RA_propCheck


--[RA_projText
--Needs: RA_lead(#myProp,#myLead), RA_revs(#myRevInfo)
--DROP TABLE #myRevInfo, #mySumm
SELECT lead, upld_date, convert(text, convert(varchar(16384),js.PROJ_SUMM_TXT) + convert(varchar(16384),js.INTUL_MERT) + convert(varchar(16384),js.BRODR_IMPT)) as summ
INTO #mySumm FROM #myLead p 
JOIN csd.prop_subm_ctl_vw psc ON psc.prop_id = p.lead
JOIN FLflpdb.flp.proj_summ js ON js.SPCL_CHAR_PDF <> 'Y' AND js.TEMP_PROP_ID = psc.TEMP_PROP_ID

SELECT p.lead, p.L, rs.string, rs.score, rp.revr_id, rp.rev_rtrn_date, id=identity(18), 0 as 'seq'
INTO #myRevInfo FROM #myLead p
JOIN csd.rev_prop_vw rp ON rp.prop_id = p.lead
JOIN tempdb.guest.revScores rs ON rs.yn = rp.rev_prop_rtng_ind 
ORDER BY lead, score DESC
SELECT lead, MIN(id) as 'start' INTO #myRISt FROM #myRevInfo GROUP BY lead
UPDATE #myRevInfo set seq = id-M.start FROM #myRevInfo r, #myRISt M WHERE r.lead = M.lead
DROP TABLE #myRISt

SELECT rv.lead, rv.L as pi_last_name, 'Review' as docType, rev_prop.pm_logn_id, rv.revr_id as panl_revr_id,  revr.revr_last_name as 'name', revr_opt_addr_line.revr_addr_txt as 'info',
convert(varchar,rev_prop.rev_type_code) AS type, convert(varchar,rev_prop.rev_stts_code) as stts, rev_prop.rev_due_date as due, rv.rev_rtrn_date as returned,
 rv.string as score, rev_txt.REV_PROP_TXT_FLDS as 'text'
FROM #myRevInfo rv, csd.revr revr, csd.rev_prop rev_prop, csd.revr_opt_addr_line revr_opt_addr_line, csd.rev_prop_txt_flds_vw rev_txt
WHERE rv.lead = rev_prop.prop_id AND rv.lead = rev_txt.PROP_ID AND rv.revr_id = rev_prop.revr_id AND rv.revr_id = revr.revr_id AND rv.revr_id = revr_opt_addr_line.revr_id
  AND rv.revr_id = rev_txt.REVR_ID AND ((revr_opt_addr_line.addr_lne_type_code='E'))
UNION ALL SELECT p.lead, p.L, ' PanlSumm' as docType, p.PO, panl_prop_summ.PANL_ID,  ' '+panl.panl_name, panl_rcom_def.RCOM_TXT,
convert(varchar,panl_prop_summ.RCOM_SEQ_NUM), convert(varchar,panl_prop_summ.PROP_ORDR), panl.panl_bgn_date, panl_prop_summ.last_updt_tmsp,
panl_rcom_def.RCOM_ABBR, panl_prop_summ.PANL_SUMM_TXT
    FROM #myLead p 
    JOIN FLflpdb.flp.panl_prop_summ panl_prop_summ ON panl_prop_summ.PROP_ID = p.lead
    JOIN csd.panl panl ON panl.panl_id = panl_prop_summ.PANL_ID
    LEFT JOIN FLflpdb.flp.panl_rcom_def panl_rcom_def ON panl_rcom_def.RCOM_SEQ_NUM = panl_prop_summ.RCOM_SEQ_NUM AND panl_rcom_def.PANL_ID = panl_prop_summ.PANL_ID

UNION ALL SELECT p.lead, p.L, 'POCmnt', p.PO, p.prop_id, cmt.cmnt_cre_id, '', '', convert(varchar,cmt.cmnt_prop_stts_code), cmt.beg_eff_date, cmt.end_eff_date,'', cmt.cmnt
FROM #myProp p, FLflpdb.flp.cmnt_prop cmt WHERE p.prop_id = cmt.prop_id AND (p.ILN < 'M' OR LEN(cmt.cmnt) <> (SELECT LEN(l.cmnt) FROM FLflpdb.flp.cmnt_prop l WHERE p.lead = l.prop_id))
UNION ALL SELECT p.lead, p.L, 'RA' as docType, p.PO, '', ra.last_updt_user, '', '', null, null, ra.last_updt_tmsp, '', ra.prop_rev_anly_txt
FROM #myLead p, csd.prop_rev_anly_vw ra WHERE p.lead = ra.prop_id
UNION ALL SELECT p.lead, p.L, 'Abstr', p.PO, '', a.last_updt_user, a.cent_mrkr_prop, a.cent_mrkr_awd, null, null, a.last_updt_tmsp,'', a.abst_narr_txt
FROM #myProp p, csd.abst a WHERE p.prop_id = a.awd_id AND (p.ILN < 'M' OR LEN(a.abst_narr_txt) <>  (SELECT LEN(l.abst_narr_txt) FROM csd.abst l WHERE p.lead = l.awd_id))
UNION ALL SELECT p.lead, p.L, 'SummProj',p.PO, '', '', '', null, null, null, s.upld_date,'', s.summ
FROM #myLead p, #mySumm s WHERE s.lead = p.lead
UNION ALL SELECT p.lead, p.L, 'xDiaryNt',p.PO, p.prop_id, crtd_by_user, ej_diry_note_kywd, null, null, null, crtd_date,'', ej_diry_note_txt
FROM #myProp p, FLflpdb.flp.ej_diry_note d WHERE d.prop_id = p.prop_id
ORDER BY lead, docType, revr_last_name, revr_id
--]RA_projText

--consider adding scribe?  I'm having timing issues with these
--select pav.panl_id +'('+rtrim(revr.revr_last_name)+')', pav.* from FLflpdb.flp.panl_asgn_view pav JOIN csd.revr revr ON revr.revr_id = pav.Scribe WHERE panl_id = 'p180637'
--select p.* from #myLead p

--SELECT prop_id, pav.panl_id, '('+rtrim(revr.revr_last_name)+')' as scribe
--FROM #myLead p
--JOIN FLflpdb.flp.panl_asgn_view pav ON pav.prop_id 
--JOIN csd.revr revr ON revr.revr_id = pav.Scribe 
--WHERE panl_id = 'p180637'



--splits details
--[RA_splits
--Needs: RA_lead(#myProp), RA_propPRCs, RA_prop(#myPropInfo)
--DROP TABLE: #myBSprc
SELECT prop_id,splt_id,budg_yr,R,id=identity(18), 0 as 'seq'
INTO #myBSprc FROM (SELECT DISTINCT p.prop_id,splt_id,budg_yr,bpr.pgm_ref_code AS R
      FROM #myProp p
      JOIN csd.budg_pgm_ref bpr ON bpr.prop_id = p.prop_id) bpr
ORDER BY prop_id,splt_id,budg_yr,R
SELECT prop_id,splt_id,budg_yr, MIN(id) as 'start' INTO #myBSst FROM #myBSprc 
GROUP BY prop_id,splt_id,budg_yr
UPDATE #myBSprc set seq = id-M.start FROM #myBSprc rb, #myBSst M
WHERE rb.prop_id = M.prop_id AND rb.splt_id=M.splt_id AND rb.budg_yr=M.budg_yr
CREATE INDEX myBSprc ON #myBSprc(prop_id,splt_id,budg_yr, seq)
DROP TABLE #myBSst

SELECT nullif(dd_rcom_date,'1900-01-01') AS dd_rcom_date, nsf_rcvd_date,  
cntx_stmt_id, prop.pgm_annc_id, prop.org_code as pOrg, bs.org_code as bOrg, 
prop.pgm_ele_code as pPEC,bs.pgm_ele_code as bPEC,p.PO as pPO, bs.pm_ibm_logn_id as bPO,
prop_stts_abbr,natr_rqst.natr_rqst_abbr,prop.obj_clas_code as pObj, bs.obj_clas_code as bObj,
ai.awd_istr_abbr as rcom_istr_abbr, p.lead,p.ILN, org.dir_div_abbr as Div, p.prop_id,
p.L AS pi_last_name, p.F AS pi_frst_name, p.I AS inst, p.M as email,
p.R as propPRCs, 
(SELECT MAX( CASE bp.seq WHEN 0 THEN rtrim(bp.R) END)+
    MAX( CASE bp.seq WHEN 1 THEN ' ' + rtrim(bp.R) END)+
    MAX( CASE bp.seq WHEN 2 THEN ' ' + rtrim(bp.R) END)+
    MAX( CASE bp.seq WHEN 3 THEN ' ' + rtrim(bp.R) END)+
    MAX( CASE bp.seq WHEN 4 THEN ' ' + rtrim(bp.R) END)+
    MAX( CASE bp.seq WHEN 5 THEN ' ' + rtrim(bp.R) END)+
    MAX( CASE bp.seq WHEN 6 THEN ' ' + rtrim(bp.R) END)+
    MAX( CASE bp.seq WHEN 7 THEN ' ' + rtrim(bp.R) END)+
    MAX( CASE bp.seq WHEN 8 THEN ' ' + rtrim(bp.R) END)+
    MAX( CASE bp.seq WHEN 9 THEN ' ' + rtrim(bp.R) END)+
    MAX( CASE bp.seq WHEN 10 THEN ' ' + rtrim(bp.R) END) 
    FROM #myBSprc bp
    WHERE bp.prop_id = bs.prop_id AND bp.budg_yr = bs.budg_yr 
          AND bp.splt_id = bs.splt_id ) AS bPRCs,
bs.awd_id, bs.budg_yr, bs.splt_id, bs.budg_splt_tot_dol,
prop_stts_txt, prop.prop_titl_txt
FROM #myPropInfo p
JOIN csd.prop prop ON prop.prop_id = p.prop_id
JOIN FLflpdb.flp.org org ON org.org_code = prop.org_code
JOIN csd.prop_stts prop_stts ON prop_stts.prop_stts_code = prop.prop_stts_code
JOIN csd.natr_rqst natr_rqst ON natr_rqst.natr_rqst_code = prop.natr_rqst_code
JOIN csd.budg_splt bs on bs.prop_id=p.prop_id 
JOIN csd.awd_istr ai on prop.rcom_awd_istr = ai.awd_istr_code
ORDER BY p.lead, p.ILN, p.prop_id, bs.budg_yr, bs.splt_id
--]RA_splits


-- Country, cover & awd check info
--FOR_PROF_FLAG,
--SMAL_BUS_FLAG,
--MINR_BUS_FLAG,
--WMEM_OWN_FLAG,
--PREV_AWD_ID,
--PREV_FL_AWD_ID,
--SUPP_FLAG,
--RNEW_FLAG,
--ACBR_FLAG,
-- does budget PRCs by award, combining all splits
--[RA_awdCheck
-- Needs: RA_lead(#myProp),RA_prop(#myPropInfo,#myPropBudg),RA_revs(#myRevs, #myRevMarks, #myRevPanl), RA_panl(#myPanl, #myProjPanl, #myProjPanlSumm)
--DROP TABLE #myCtry, #myCovrInfo 
SELECT p.prop_id, ctry.ctry_name, id=identity(18), 0 as 'seq' 
INTO #myCtry FROM #myProp p
JOIN csd.prop_subm_ctl_vw psc ON psc.prop_id = p.prop_id
JOIN csd.prop_spcl_item_vw sp1 ON sp1.TEMP_PROP_ID = psc.TEMP_PROP_ID
JOIN csd.ctry ctry ON sp1.SPCL_ITEM_CODE = ctry.ctry_code
WHERE end_date Is Null
ORDER BY p.prop_id, ctry.ctry_name
CREATE INDEX myCtry_idx ON #myCtry(prop_id) 
SELECT prop_id, MIN(id) as 'start' INTO #myStCtry FROM #myCtry GROUP BY prop_id
UPDATE #myCtry set seq = id-M.start FROM #myCtry r, #myStCtry M WHERE r.prop_id = M.prop_id
DROP TABLE #myStCtry

SELECT p.prop_id,OTH_AGCY_SUBM_FLAG,
CASE WHEN PC.HUM_DATE is not NULL THEN convert(varchar(10),PC.HUM_DATE,1) WHEN PC.humn_date_pend_flag='Y' THEN 'Pend' END AS humn_date,
humn_subj_asur_num AS hum_asur, HUM_EXPT,
CASE WHEN PC.VERT_DATE is not NULL THEN convert(varchar(10),PC.VERT_DATE,1) WHEN PC.vrtb_date_pend_flag='Y' THEN 'Pend' END AS vrtb_date,
vrtb_wlfr_asur_num AS vrtb_asur
INTO #myCovrInfo FROM #myProp p
JOIN csd.prop_subm_ctl_vw psc ON psc.prop_id = p.prop_id
JOIN FLflpdb.flp.PROP_COVR PC ON PC.TEMP_PROP_ID = psc.TEMP_PROP_ID
CREATE INDEX myCovrInfo_ix ON #myCovrInfo(prop_id)

-- checking those recommended for award
SELECT nullif(dd_rcom_date,'1900-01-01') AS dd_rcom_date, nsf_rcvd_date, 
prop.pgm_annc_id, prop.org_code, prop.pgm_ele_code, 
prop_stts_abbr,natr_rqst.natr_rqst_abbr,prop.obj_clas_code,ai.awd_istr_abbr as rcom_istr_abbr,
cntx_stmt_id, p.PO, ra.RAupdate, p.lead, p.ILN, --eJ link
org.dir_div_abbr as Div, p.prop_id, 
p.L AS pi_last_name, p.F AS pi_frst_name, p.I AS inst, p.st_code, 
c.humn_date, c.hum_asur, c.HUM_EXPT, c.vrtb_date, c.vrtb_asur, c.OTH_AGCY_SUBM_FLAG, 
(SELECT MAX(CASE y.seq WHEN 1 THEN y.ctry_name ELSE '' END)+
    MAX(CASE y.seq WHEN 2 THEN '; '+y.ctry_name ELSE '' END)+
    MAX(CASE y.seq WHEN 3 THEN '; '+y.ctry_name ELSE '' END)+
    MAX(CASE y.seq WHEN 4 THEN '; '+y.ctry_name ELSE '' END)+
    MAX(CASE y.seq WHEN 5 THEN '; '+y.ctry_name ELSE '' END)+
    MAX(CASE y.seq WHEN 6 THEN '; '+y.ctry_name ELSE '' END)+
    MAX(CASE y.seq WHEN 7 THEN '; '+y.ctry_name ELSE '' END)+
    MAX(CASE y.seq WHEN 8 THEN '; '+y.ctry_name ELSE '' END) 
    FROM #myCtry y WHERE y.prop_id = p.prop_id) AS Country,
b.frgn_trav_tot, a.abst_narr_txt, prop_stts_txt, prop.prop_titl_txt,
prop.rqst_eff_date,prop.rqst_mnth_cnt,
b.RN as brevn,p.D AS rqst_tot, b.T as budg_tot, 
(SELECT budg_tot_dol FROM #myBudg e WHERE e.prop_id = p.prop_id AND budg_seq_yr = 1) as y1_tot,
nullif((SELECT budg_tot_dol FROM #myBudg e WHERE e.prop_id = p.prop_id AND budg_seq_yr = 2),0) as y2_tot,
nullif((SELECT budg_tot_dol FROM #myBudg e WHERE e.prop_id = p.prop_id AND budg_seq_yr = 3),0) as y3_tot,
nullif((SELECT budg_tot_dol FROM #myBudg e WHERE e.prop_id = p.prop_id AND budg_seq_yr = 4),0) as y4_tot,
nullif((SELECT budg_tot_dol FROM #myBudg e WHERE e.prop_id = p.prop_id AND budg_seq_yr = 5),0) as y5_tot,
nullif((SELECT budg_tot_dol FROM #myBudg e WHERE e.prop_id = p.prop_id AND budg_seq_yr = 6),0) as y6_tot,
b.sub_ctr_tot, b.pdoc_tot,b.part_tot_dol, b.grad_tot_cnt, 
b.sr_tot_cnt, b.sr_sumr_mnths, b.sr_acad_mnths, b.sr_cal_mnths, 
p.M as email
FROM #myPropInfo p
JOIN csd.prop prop ON prop.prop_id = p.prop_id
JOIN FLflpdb.flp.org org ON org.org_code = prop.org_code
JOIN csd.prop_stts prop_stts ON prop_stts.prop_stts_code = prop.prop_stts_code
JOIN csd.natr_rqst natr_rqst ON natr_rqst.natr_rqst_code = prop.natr_rqst_code
LEFT JOIN csd.awd_istr ai on prop.rcom_awd_istr = ai.awd_istr_code
LEFT JOIN #myRA ra ON ra.lead = p.lead
LEFT JOIN csd.abst a ON a.awd_id = p.prop_id
LEFT JOIN #myPropBudg b ON b.prop_id = p.prop_id
LEFT JOIN #myCovrInfo c ON c.prop_id = p.prop_id
ORDER BY p.lead, p.ILN, p.prop_id
--]RA_awdCheck

--Get budget details
--[RA_budgBlocks
--Needs: RA_lead(#myProp), RA_propPRCs(#myPRCs), RA_budg(#mBudg)

SELECT p.nsf_rcvd_date, p.pgm_annc_id, p.org_code, p.pgm_ele_code, 
pid.PO, pid.lead, pid.ILN, pid.prop_id, '' as pi_last_name, '' as inst_name, '' as pi_email, 
eb.sr_pers_cnt, eb.sr_summ_mnth_cnt, eb.pdoc_grnt_dol, eb.frgn_trav_dol, eb.part_dol, eb.budg_tot_dol, 
eb.revn_num, eb.budg_seq_yr, eb.budg_tot_dol AS dol, p.org_code, p.pgm_ele_code as PEC, p.obj_clas_code as Obj, prc.R as PRCs
FROM #myProp pid
JOIN csd.prop p ON p.prop_id = pid.prop_id
LEFT OUTER JOIN #myPRCs prc ON prc.prop_id = pid.prop_id
LEFT OUTER JOIN #myBudg eb ON eb.prop_id = pid.prop_id 
UNION ALL SELECT p.nsf_rcvd_date, p.pgm_annc_id, p.org_code, p.pgm_ele_code, 
pid.PO, pid.lead, pid.ILN, pid.prop_id, pid.L as pi_last_name, pid.I as inst_name, pid.M as pi_email, 
null, null, null, null, null, null, (SELECT MAX(revn_num) from #myBudg b WHERE b.prop_id = pid.prop_id),
null, (SELECT SUM(budg_tot_dol) from #myBudg b WHERE b.prop_id = pid.prop_id), p.org_code, p.pgm_ele_code, p.obj_clas_code, prc.R 
FROM #myProp pid
JOIN csd.prop p ON p.prop_id = pid.prop_id
LEFT OUTER JOIN #myPRCs prc ON prc.prop_id = pid.prop_id
ORDER BY pid.lead, pid.ILN, pid.prop_id, eb.budg_seq_yr 
--]RA_budgBlocks
--select * from #myPRCs

--table permission checking
-- These are all the tables we use (or would think of using.)
SELECT 'abst' AS tbl INTO #myTbl
union all SELECT 'addl_pi_invl'
union all SELECT 'awd_istr'
union all SELECT 'budg_pgm_ref'
union all SELECT 'budg_splt'
union all SELECT 'ctry'
union all SELECT 'ej_upld_doc_vw'
union all SELECT 'eps_blip'
union all SELECT 'inst'
union all SELECT 'natr_rqst'
union all SELECT 'org' -- use flp
union all SELECT 'panl'
union all SELECT 'panl_prop'
union all SELECT 'panl_revr'
union all SELECT 'pgm_annc'
union all SELECT 'pgm_ele' -- use flp
union all SELECT 'pgm_ref' -- only non-public table
union all SELECT 'PI_dmog'
union all SELECT 'pi_vw'
union all SELECT 'po_vw'
union all SELECT 'prop'
union all SELECT 'prop_atr'
union all SELECT 'prop_rev_anly_vw'
union all SELECT 'prop_spcl_item_vw'
union all SELECT 'prop_stts'
union all SELECT 'prop_subm_ctl_vw'
union all SELECT 'rev_prop'
union all SELECT 'rev_prop_txt_flds_vw'
union all SELECT 'rev_prop_vw'
union all SELECT 'revr'
union all SELECT 'revr_opt_addr_line'
union all SELECT 'cmnt_prop' -- begin flp tables
union all SELECT 'ej_diry_note'
union all SELECT 'obj_clas_pars' -- use flp
union all SELECT 'panl_asgn_view'
union all SELECT 'panl_prop_summ'
union all SELECT 'panl_rcom_def'
union all SELECT 'proj_summ'
union all SELECT 'PROP_COVR'
union all SELECT 'obj_clas' -- don't use
union all SELECT 'pgm_ele_pars' -- don't need
union all SELECT 'pgm_ref_pars' -- doesn't exist
union all SELECT 'org_pars' -- don't need 
order by tbl

-- tables with not public read access in rptdb or FLflpdb
-- the conclusion: use flp.obj_clas_pars, flp.org, flp.pgm_ele
select t.* 
FROM (SELECT t.* from #myTbl t
where not exists (SELECT * FROM rptdb.dbo.sysobjects so 
JOIN rptdb.dbo.sysprotects sp ON sp.id = so.id and sp.action = 193 and sp.uid = 0
WHERE so.name = t.tbl)) t
where not exists (SELECT * FROM FLflpdb.dbo.sysobjects so 
JOIN FLflpdb.dbo.sysprotects sp ON sp.id = so.id and sp.action = 193 and sp.uid = 0
WHERE so.name = t.tbl)

select dbu.name as db, t.tbl, so.crdate, so.type, su.name as username
from #myTbl t
JOIN FLflpdb.dbo.sysobjects so ON so.name = t.tbl
JOIN FLflpdb.dbo.sysusers dbu ON dbu.uid = so.uid
JOIN FLflpdb.dbo.sysprotects sp ON sp.id = so.id and sp.action = 193 
JOIN FLflpdb.dbo.sysusers su ON su.uid = sp.uid
union all
select dbu.name as db, t.tbl, so.crdate, so.type, su.name as username
from #myTbl t
JOIN rptdb.dbo.sysobjects so ON so.name = t.tbl
JOIN rptdb.dbo.sysusers dbu ON dbu.uid = so.uid
JOIN rptdb.dbo.sysprotects sp ON sp.id = so.id and sp.action = 193 
JOIN rptdb.dbo.sysusers su ON su.uid = sp.uid
order by tbl, username, db

Drop table #myTbl

DROP TABLE #myPidRAt
DROP TABLE #myProp, #myLead, #myRA 
DROP TABLE #myPRCs, #myPRCdata
DROP TABLE #myBudg
DROP TABLE #myPropBudg, #myPropInfo
DROP TABLE #myRevs, #myRevPanl, #myRevMarks, #myRevSumm
DROP TABLE #myPPConfl
DROP TABLE #myPanl, #myProjPanl, #myProjPanlSumm
DROP TABLE #myRevInfo, #mySumm
DROP TABLE #myDmog
DROP TABLE #myCtry, #myCovrInfo
DROP TABLE #myBSprc


