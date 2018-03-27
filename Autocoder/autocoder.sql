SET NOCOUNT ON
SELECT lead_prop_id, prop_id, pi_id INTO #myProp FROM csd.prop prop WHERE 
(prop.prop_id  in ('1810758','1812466','1812695','1812746','1812766','1812927','1812944','1812950','1813165','1813303','1813511','1813624','1813782','1813950','1814041','1814333','1814432','1814511','1814613','1814886','1815029','1815050','1815190','1815316','1815383','1815463','1815709','1815819','1815846','1815885','1815980','1816094','1816621','1816658','1816922','1817212','1812698','1813357','1813359','1813957','1814662','1814873','1814892','1814925','1815439','1815745','1815842','1815870','1815898','1816033','1816329','1816501','1816797','1817222') OR prop.lead_prop_id  in ('1810758','1812466','1812695','1812746','1812766','1812927','1812944','1812950','1813165','1813303','1813511','1813624','1813782','1813950','1814041','1814333','1814432','1814511','1814613','1814886','1815029','1815050','1815190','1815316','1815383','1815463','1815709','1815819','1815846','1815885','1815980','1816094','1816621','1816658','1816922','1817212','1812698','1813357','1813359','1813957','1814662','1814873','1814892','1814925','1815439','1815745','1815842','1815870','1815898','1816033','1816329','1816501','1816797','1817222'))
--[AC_check
CREATE INDEX myProp_ix ON #myProp(prop_id)
INSERT #myProp SELECT prop.lead_prop_id, prop.prop_id, prop.pi_id FROM csd.prop prop
JOIN (SELECT DISTINCT lead_prop_id FROM #myProp WHERE NOT lead_prop_id IS NULL) leads ON leads.lead_prop_id = prop.lead_prop_id 
WHERE NOT EXISTS (SELECT * FROM #myProp p WHERE p.prop_id = prop.prop_id)
--select * from #myProp order by lead_prop_id, prop_id

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

SELECT DISTINCT prop.prop_id, pa.prop_atr_code, id=identity(18), 0 as 'seq' INTO #myPRCs
FROM #myProp prop, csd.prop_atr pa WHERE pa.prop_id = prop.prop_id  AND pa.prop_atr_type_code = 'PRC'
ORDER BY prop.prop_id, pa.prop_atr_code
CREATE INDEX myPRCs_idx ON #myPRCs(prop_id)
SELECT prop_id, MIN(id) as 'start' INTO #myStarts FROM #myPRCs GROUP BY prop_id
UPDATE #myPRCs set seq = id-M.start FROM #myPRCs r, #myStarts M WHERE r.prop_id = M.prop_id

SELECT isnull(prop.lead_prop_id,prop.prop_id) AS lead, 
CASE WHEN prop.lead_prop_id IS NULL THEN 'I' WHEN prop.lead_prop_id <> prop.prop_id THEN 'N' ELSE 'L' END AS ILN, 
prop.prop_id, cntx_stmt_id, pi_last_name, pi_frst_name, inst.inst_shrt_name, inst.st_code,
prop_stts_abbr, pm_ibm_logn_id, 
bas_rsch_pct, apld_rsch_pct, educ_trng_pct, mjr_equp_pct, land_buld_fix_equp_pct, non_invt_pct,
prop.org_code, CASE WHEN prop.org_code<>prop.orig_org_code THEN prop.orig_org_code END AS orig_org,
prop.pgm_ele_code, CASE WHEN prop.pgm_ele_code <> prop.orig_pgm_ele_code THEN prop.orig_pgm_ele_code END AS orig_PEC, 
NfmlPIs,NhispPIs,NhndcpPIs,NnonWhtAsnPIs,
(SELECT MAX( CASE pa.seq WHEN 0 THEN pa.prop_atr_code ELSE '' END ) + ' ' +
        MAX( CASE pa.seq WHEN 1 THEN pa.prop_atr_code ELSE '' END ) + ' ' +
        MAX( CASE pa.seq WHEN 2 THEN pa.prop_atr_code ELSE '' END ) + ' ' +
        MAX( CASE pa.seq WHEN 3 THEN pa.prop_atr_code ELSE '' END ) + ' ' +
        MAX( CASE pa.seq WHEN 4 THEN pa.prop_atr_code ELSE '' END ) + ' ' +
        MAX( CASE pa.seq WHEN 5 THEN pa.prop_atr_code ELSE '' END ) + ' ' +
        MAX( CASE pa.seq WHEN 6 THEN pa.prop_atr_code ELSE '' END ) FROM #myPRCs pa WHERE p.prop_id = pa.prop_id) AS PRCs,
pgm_annc_id, prop_titl_txt
FROM #myProp p, #myDmog d, csd.inst inst, csd.pi_vw pi_vw, csd.prop prop, csd.prop_stts prop_stts
WHERE prop.pi_id = pi_vw.pi_id AND prop.inst_id = inst.inst_id AND p.prop_id = prop.prop_id 
AND d.prop_id = p.prop_id AND prop.prop_stts_code = prop_stts.prop_stts_code
ORDER BY lead, ILN, prop.prop_id
DROP TABLE #myStarts DROP TABLE #myPRCs DROP TABLE #myProp DROP TABLE #myDmog
--]AC_check