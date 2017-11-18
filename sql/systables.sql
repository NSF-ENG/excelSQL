-- table & column names from fields 
select so.name, sc.name, t.name, sc.length 
from syscolumns sc 
JOIN sysobjects so ON sc.id = so.id
JOIN systypes t on t.usertype = sc.usertype
where sc.name like '%ibm_logn%'
order by so.name, sc.name

select so.name, sc.name, t.name, sc.length, su.name, sp.uid, sp.action, sp.protecttype, sp.columns
from syscolumns sc 
JOIN sysobjects so ON so.id = sc.id
JOIN systypes t on t.usertype = sc.usertype
JOIN sysprotects sp ON sp.id = sc.id
JOIN sysusers su ON su.uid = sp.uid
where sc.name like '%logn%' -- field name
AND sp.uid = 0 -- skip to see non-public tables, too
order by so.name, sc.name, su.name

-- column names & types for table
SELECT sc.colid, sc.name, t.name, sc.length FROM sysobjects so 
JOIN syscolumns sc ON sc.id = so.id 
JOIN systypes t on t.usertype = sc.usertype
WHERE so.name like '#myTmp%' -- table name
ORDER BY colid

select so.name, so.type, so.id from sysobjects so where so.name like '%myTmp%'

select so.name, sc.name,t.name, sc.length 
from syscolumns sc, sysobjects so, systypes t
where sc.id = so.id and t.usertype = sc.usertype and so.name in ('prop_fl_budg_yr_map','intl_impl','grnt_note'
)
order by so.name, sc.name

select so.name, so.id, so.uid, sp.*
FROM FLflpdb.dbo.sysobjects so, FLflpdb.dbo.sysprotects sp
WHERE so.name = 'clbr_affl' and sp.id = so.id

select so.name, so.id, so.uid, su.name, sp.*  from sysprotects sp, sysusers su, sysobjects so
where sp.uid = su.uid and sp.id = so.id and so.name = 'INST' --su.name in ('ccfuser','aciuser','cnsuser','iisuser')
-- sp.uid = su.uid and sp.id = so.id and so.name = 'tz_ctry_map' 
order by so.name, su.name

select so.name, sc.name,t.name, sc.length 
from syscolumns sc, sysobjects so, systypes t
where sc.id = so.id and t.usertype = sc.usertype and so.name = 'revr_opt_addr_line'
order by so.name, sc.name

select count(*) from flp.snap_dmog -- 66,420 
select count(*) from flp.psnap_dmog -- 66,420 
select count(*) from flp.ppi_dmog -- 402,914
select count(*) from flp.pi_dmog -- 402,914 
select top 50 * from flp.psnap_dmog where prop_id like '1%'
select count(*) from flp.pi -- 601,000

--select top 20 * from csd.revr_appt where last_updt_tmsp > '1/1/17'
--select top 20 * from csd.revr_attr where last_updt_tmsp > '1/1/17'
--select top 20 * from csd.rev_prop_sgst where last_updt_tmsp > '1/1/17'
--select top 20 * from csd.revr_addl_fos where last_updt_tmsp > '1/1/17'
--select top 20 * from csd.revr_narr where last_updt_tmsp > '1/1/17'
--select top 20 * from csd.revr_salu where last_updt_tmsp > '1/1/17'


