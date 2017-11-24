-- Given a list of projects, look up all PI & co-PI current awards and pending proposals.

SET NOCOUNT ON 
--[proppiProps
SELECT isnull(lead_prop_id,prop.prop_id) AS 'lead', prop.prop_id, prop.pi_id 
INTO #myProps
FROM csd.prop prop
--]proppiProps
WHERE prop.prop_id In ('1111111','1234567','1749678','1749385','1750127','1750076','1751047','1749563','1751254','1749414','1749734') 
OR prop.lead_prop_id In ('1111111','1234567','1749678','1749385','1750127','1750076','1751047','1749563','1751254','1749414','1749734') 
--[proppiAwd
SELECT PIs.lead, PIs.prop_id, PIs.pi_id, awd.awd_id, awd_pi_copi.proj_role_code
INTO #myAwds 
FROM (SELECT * FROM #myProps
 UNION SELECT prop.lead, prop.prop_id, addl_pi_invl.pi_id
 FROM csd.addl_pi_invl addl_pi_invl, #myProps prop
WHERE prop.prop_id = addl_pi_invl.prop_id) PIs, 
csd.awd awd, csd.awd_pi_copi awd_pi_copi 
WHERE awd_pi_copi.awd_id = awd.awd_id  AND (awd.awd_stts_code='80') AND (PIs.pi_id = awd_pi_copi.pi_id) AND awd_pi_copi.end_date IS NULL 
SELECT DISTINCT #myAwds.awd_id INTO #myDistAwds FROM #myAwds
SELECT a.lead, a.prop_id, pi_vw.pi_last_name, pi_vw.pi_frst_name, pi_vw.pi_emai_addr, inst.inst_name, a.pi_id, a.proj_role_code, a.awd_id, instaw.inst_name, awd.tot_intn_awd_amt, tPI.totPIs,  ob.totOblg_amt, ob.totCuml_xpnd_amt, ob.last_pymt_date, awd.last_prop_id, awd.org_code, awd.pgm_ele_code, awd.pm_ibm_logn_id, awd.awd_stts_date, awd.awd_eff_date, awd.awd_exp_date,  awd.awd_titl_txt 
FROM #myAwds a, 
 (SELECT awd1.awd_id, Count(DISTINCT awd_pi_copi1.pi_id) AS 'totPIs'
  FROM #myDistAwds awd1, csd.awd_pi_copi awd_pi_copi1 
  WHERE awd1.awd_id = awd_pi_copi1.awd_id AND (awd_pi_copi1.end_date Is Null) 
  GROUP BY awd1.awd_id) tPI, 
 (SELECT awd2.awd_id, Sum(oblg_bal.oblg_amt) AS 'totOblg_amt', Sum(oblg_bal.cuml_xpnd_amt) AS 'totCuml_xpnd_amt', Max(oblg_bal.last_pymt_date) AS 'last_pymt_date' 
  FROM #myDistAwds awd2, csd.oblg_bal oblg_bal 
  WHERE awd2.awd_id = oblg_bal.oblg_id  
  GROUP BY awd2.awd_id) ob, 
csd.awd awd, csd.inst inst, csd.pi_vw pi_vw, csd.inst instaw 
WHERE a.awd_id=awd.awd_id AND a.awd_id=ob.awd_id AND a.awd_id=tPI.awd_id AND a.pi_id = pi_vw.pi_id AND awd.inst_id = instaw.inst_id  AND pi_vw.inst_id = inst.inst_id 
ORDER BY a.lead, a.prop_id, pi_vw.pi_last_name, pi_vw.pi_frst_name, a.pi_id, awd.awd_exp_date 
DROP TABLE #myAwds,#myDistAwds 
--DROP TABLE #myProps
--]proppiAwd

--[proppiPend
SELECT * INTO #myPIs FROM #myProps
 UNION SELECT prop.lead, prop.prop_id, addl_pi_invl.pi_id 
 FROM csd.addl_pi_invl addl_pi_invl, #myProps prop
 WHERE prop.prop_id = addl_pi_invl.prop_id
SELECT PIs.lead, PIs.prop_id, PIs.pi_id, '1' as 'proj_role_code', prop.prop_id AS pend_id
INTO #myPend FROM #myPIs PIs, csd.prop prop
WHERE (prop.pi_id = PIs.pi_id) AND prop.prop_stts_code IN ('00','01','02','03','06','40') AND prop.prop_id <> PIs.prop_id
UNION SELECT PIs.lead, PIs.prop_id, PIs.pi_id, addl_pi_invl.proj_role_code, addl_pi_invl.prop_id AS pend_id
FROM #myPIs PIs, csd.prop prop, csd.addl_pi_invl addl_pi_invl
WHERE addl_pi_invl.pi_id = PIs.pi_id AND addl_pi_invl.prop_id = prop.prop_id AND prop.prop_stts_code IN ('00','01','02','03','06','40') AND addl_pi_invl.prop_id <> PIs.prop_id
SELECT a.lead, a.prop_id, pi_vw.pi_last_name, pi_vw.pi_frst_name, pi_vw.pi_emai_addr, inst.inst_name, a.pi_id, a.proj_role_code,
a.pend_id, prop_stts.prop_stts_txt, pps.panl_id, pr.rcom_seq_num, pr.rcom_abbr,
prop.rqst_dol, prop.rqst_mnth_cnt, nullif(prop.rcom_mnth_cnt,0) AS 'rcom_mnth_cnt', nullif(prop.rcom_awd_istr,'0') AS 'rcom_awd_istr',
prop.rqst_eff_date, nullif(prop.rcom_eff_date,'1900-01-01') AS 'rcom_eff_date', nullif(prop.pm_asgn_date,'1900-01-01') AS pm_asgn_date,
nullif(prop.pm_rcom_date,'1900-01-01') AS  pm_rcom_date,
prop.nsf_rcvd_date, prop.org_code, prop.pgm_ele_code, prop.pm_ibm_logn_id, prop.prop_stts_code, prop_stts.prop_stts_abbr,
(SELECT 1+COUNT(*) FROM csd.addl_pi_invl api  WHERE a.pend_id = api.prop_id) AS numPI, prop.prop_titl_txt
FROM #myPend a, csd.prop prop, csd.pi_vw pi_vw, csd.inst inst, csd.prop_stts prop_stts, flflpdb.flp.panl_prop_summ pps, flflpdb.flp.panl_rcom_def pr
WHERE a.pi_id = pi_vw.pi_id AND a.pend_id = prop.prop_id AND pi_vw.prim_addr_flag = 'Y' AND prop.inst_id = inst.inst_id AND prop.prop_stts_code = prop_stts.prop_stts_code
 AND a.pend_id *= pps.prop_id AND pps.panl_id *= pr.panl_id AND pps.rcom_seq_num *= pr.rcom_seq_num
ORDER BY a.lead, a.prop_id, pi_vw.pi_last_name, pi_vw.pi_frst_name, a.pi_id, prop.rqst_dol DESC
DROP TABLE #myPend, #myProps,#myPIs
--]proppiPend