SET NOCOUNT ON 
SELECT CASE WHEN p.lead_prop_id IS NULL THEN 'I' WHEN p.lead_prop_id <> p.prop_id THEN 'N' ELSE 'L' END AS ILN,
isnull(p.lead_prop_id,p.prop_id) AS lead, p.prop_id, psc.TEMP_PROP_ID
INTO #myPids FROM csd.prop p
JOIN csd.prop_subm_ctl_vw psc ON psc.prop_id = p.prop_id
JOIN csd.panl_prop pp ON p.prop_id = pp.prop_id
WHERE p.prop_stts_code IN ('00','01','02','08','09') AND pp.panl_id in ('p180207')
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


-- PRC
SELECT DISTINCT p.prop_id, pa.prop_atr_code, id=identity(18), 0 as 'seq' INTO #myPRCs
FROM #myPids p
JOIN csd.prop_atr pa ON pa.prop_id = p.prop_id AND pa.prop_atr_type_code = 'PRC'
ORDER BY p.prop_id, pa.prop_atr_code
CREATE INDEX myPRCs_idx ON #myPRCs(prop_id) 
SELECT prop_id, MIN(id) as 'start' INTO #myPRCStart FROM #myPRCs GROUP BY prop_id
UPDATE #myPRCs set seq = id-M.start FROM #myPRCs r, #myPRCStart M WHERE r.prop_id = M.prop_id
DROP TABLE #myPRCStart
--dmog
SELECT prop_id, 
SUM(CASE WHEN pi_gend_code = 'F'THEN 1 ELSE 0 END) AS NfmlPIs,
SUM(CASE WHEN pi_ethn_code = 'H'THEN 1 ELSE 0 END) AS NhispPIs,
SUM(CASE WHEN dmog_tbl_code = 'H'AND dmog_code <> 'N' THEN 1 ELSE 0 END) AS NhndcpPIs,
SUM(CASE WHEN dmog_tbl_code = 'R'AND dmog_code NOT IN ('U','W','B3') THEN 1 ELSE 0 END) AS NnonWhtAsnPIs
INTO #myDmog
FROM (SELECT p.prop_id, pi_id FROM #myPids p JOIN csd.prop prop ON prop.prop_id = p.prop_id
      UNION ALL SELECT p.prop_id, pi_id FROM #myPids p JOIN csd.addl_pi_invl a ON a.prop_id = p.prop_id) PIs
LEFT OUTER JOIN csd.pi_vw pi ON pi.pi_id = PIs.pi_id
LEFT OUTER JOIN csd.PI_dmog d ON d.pi_id = PIs.pi_id
GROUP BY prop_id
ORDER BY prop_id
CREATE INDEX myDmog_idx ON #myDmog(prop_id)

-- revs
SELECT p.prop_id, COUNT(*) as Nrev,
nullif(SUM(CASE WHEN rpv.rev_prop_unrl_flag = 'Y' THEN 1 ELSE 0 END),0) as NrevUnreleasable,
nullif(SUM(CASE WHEN rpv.rev_rlse_flag = 'Y' THEN 0 WHEN rpv.rev_prop_unrl_flag = 'Y'THEN 0 ELSE 1 END),0) as NrevUnmarked 
INTO #myRevs 
FROM #myPids p
JOIN csd.rev_prop rp ON rp.prop_id = p.prop_id AND rp.rev_stts_code <> 'C'
JOIN csd.rev_prop_vw rpv ON rpv.prop_id = p.prop_id AND rpv.revr_id = rp.revr_id 
GROUP BY p.prop_id
ORDER BY p.prop_id
-- summ
SELECT p.prop_id, COUNT(pp.panl_id) as Npanl,
nullif(SUM(CASE WHEN pps.panl_summ_unrl_flag = 'Y' THEN 1 ELSE 0 END),0) as NpanlUnreleasable,
nullif(SUM(CASE WHEN pps.panl_summ_rlse_flag = 'Y' THEN 0 WHEN pps.panl_summ_unrl_flag = 'Y'THEN 0 ELSE 1 END),0) as NpanlUnmarked  
INTO #myPanl
FROM #myPids p
JOIN csd.panl_prop pp ON pp.prop_id = p.prop_id
LEFT OUTER JOIN FLflpdb.flp.panl_prop_summ pps ON pps.prop_id = p.prop_id AND pps.panl_id = pp.panl_id
GROUP BY p.prop_id
ORDER BY p.prop_id
-- props: get codes to check if they match leads
SELECT p.nsf_rcvd_date, nullif(p.dd_rcom_date,'1900-01-01') AS dd_rcom_date,
ILN, lead, p.prop_id,
pi.pi_last_name, pi.pi_frst_name, inst.inst_shrt_name AS inst_name, 
pi.pi_emai_addr,p.rqst_dol, p.rqst_eff_date,p.rqst_mnth_cnt,p.cntx_stmt_id, p.bas_rsch_pct, p.apld_rsch_pct+p.educ_trng_pct+land_buld_fix_equp_pct+mjr_equp_pct+non_invt_pct AS other_pct,
CASE WHEN PC.HUM_DATE is not NULL THEN convert(varchar(10),PC.HUM_DATE,1) WHEN PC.humn_date_pend_flag='Y' THEN 'Pend' END AS humn_date,
CASE WHEN PC.VERT_DATE is not NULL THEN convert(varchar(10),PC.VERT_DATE,1) WHEN PC.vrtb_date_pend_flag='Y' THEN 'Pend' END AS vrtb_date,
(SELECT MAX(CASE b.seq WHEN 1 THEN b.ctry_name ELSE '' END)+
    MAX(CASE b.seq WHEN 2 THEN '; '+b.ctry_name ELSE '' END)+
    MAX(CASE b.seq WHEN 3 THEN '; '+b.ctry_name ELSE '' END)+
    MAX(CASE b.seq WHEN 4 THEN '; '+b.ctry_name ELSE '' END)+
    MAX(CASE b.seq WHEN 5 THEN '; '+b.ctry_name ELSE '' END)+
    MAX(CASE b.seq WHEN 6 THEN '; '+b.ctry_name ELSE '' END)+
    MAX(CASE b.seq WHEN 7 THEN '; '+b.ctry_name ELSE '' END)+
    MAX(CASE b.seq WHEN 8 THEN '; '+b.ctry_name ELSE '' END) 
    FROM #myCtry b WHERE b.prop_id = prop.prop_id) AS Country,
a.abst_narr_txt, (SELECT SUM(frgn_trav_dol) FROM csd.eps_blip eb WHERE eb.prop_id = prop.prop_id AND NOT EXISTS (SELECT eb1.revn_num FROM csd.eps_blip eb1 WHERE eb.prop_id = eb1.prop_id AND eb.revn_num < eb1.revn_num)) as frgn_trvl_dol,
p.pgm_annc_id, p.org_code, p.pgm_ele_code, p.pm_ibm_logn_id as PO, NfmlPIs,NhispPIs,NhndcpPIs,NnonWhtAsnPIs, Nrev,NrevUnreleasable, NrevUnmarked,  Npanl, NpanlUnreleasable, NpanlUnmarked,
p.obj_clas_code, p.natr_rqst_code, p.prop_stts_code, PRCs, p.prop_titl_txt,
p.inst_id, inst.st_code, p.pi_id, ra.last_updt_tmsp as RAupdate
FROM #myPids prop
JOIN csd.prop p ON p.prop_id = prop.prop_id
JOIN #myDmog d ON d.prop_id = prop.prop_id
LEFT OUTER JOIN csd.inst inst ON inst.inst_id = p.inst_id
LEFT OUTER JOIN csd.pi_vw pi ON pi.pi_id = p.pi_id
LEFT OUTER JOIN #myRevs rv ON rv.prop_id = prop.prop_id
LEFT OUTER JOIN #myPanl pl ON pl.prop_id = prop.prop_id
LEFT OUTER JOIN csd.abst a ON a.awd_id = prop.prop_id
LEFT OUTER JOIN csd.prop_rev_anly_vw ra ON ra.prop_id = prop.prop_id
LEFT OUTER JOIN FLflpdb.flp.PROP_COVR PC ON PC.TEMP_PROP_ID = prop.TEMP_PROP_ID
LEFT OUTER JOIN (SELECT prop_id,
        MAX( CASE pa.seq WHEN 0 THEN       pa.prop_atr_code ELSE '' END ) +
        MAX( CASE pa.seq WHEN 1 THEN ' ' + pa.prop_atr_code ELSE '' END ) +
        MAX( CASE pa.seq WHEN 2 THEN ' ' + pa.prop_atr_code ELSE '' END ) +
        MAX( CASE pa.seq WHEN 3 THEN ' ' + pa.prop_atr_code ELSE '' END ) +
        MAX( CASE pa.seq WHEN 4 THEN ' ' + pa.prop_atr_code ELSE '' END ) +
        MAX( CASE pa.seq WHEN 5 THEN ' ' + pa.prop_atr_code ELSE '' END ) +
        MAX( CASE pa.seq WHEN 6 THEN ' ' + pa.prop_atr_code ELSE '' END ) AS PRCs
    FROM #myPRCs pa 
    GROUP BY prop_id) myPRCs ON myPRCs.prop_id = prop.prop_id
WHERE pi.prim_addr_flag='Y'
ORDER BY prop.lead, prop.ILN, prop.prop_id
DROP TABLE #myDmog DROP TABLE #myPids DROP TABLE #myPRCs DROP TABLE #myCtry DROP TABLE #myRevs DROP TABLE #myPanl
