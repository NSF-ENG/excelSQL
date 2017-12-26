-- table, column names & types 
SELECT so.name as tablename, sc.colid, sc.name as fieldname, t.name as typename, sc.length 
FROM sysobjects so 
JOIN syscolumns sc ON sc.id = so.id 
JOIN systypes t on t.usertype = sc.usertype
WHERE so.name like '%' and sc.name like '%' -- table name, field name
ORDER BY so.name, colid

-- with protection information, too
SELECT so.name as tablename, so.crdate, sc.colid, sc.name as fieldname, t.name as typename, sc.length, su.name as username, sp.uid, sp.action, sp.protecttype, 
inttohex(sp.columns) as col
from sysobjects so 
JOIN syscolumns sc ON sc.id = so.id
JOIN systypes t on t.usertype = sc.usertype
LEFT JOIN sysprotects sp ON sp.id = sc.id
LEFT JOIN sysusers su ON su.uid = sp.uid
WHERE so.name like '%ref%' and sc.name like '%txt%' -- table name, field name
AND (su.name is null OR su.name = 'public'OR su.name like '%ccf%')
order by so.name, colid, su.name

SELECT so.name as tablename, so.crdate, su.name as username, sp.uid, sp.action, sp.protecttype, 
inttohex(sp.columns) as col
from sysobjects so 
LEFT JOIN sysprotects sp ON sp.id = so.id
LEFT JOIN sysusers su ON su.uid = sp.uid
WHERE so.name like '%oblg%'  -- table name, field name
--AND (su.name like 'public')
order by so.name, su.name

--select * from flp.ej_elec_aprv_stts
--select top 100 * from flp.elec_sign_doc_txt

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