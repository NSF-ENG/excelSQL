-- column names & types for table
SELECT sc.colid, sc.name, t.name, sc.length FROM sysobjects so 
JOIN syscolumns sc ON sc.id = so.id 
JOIN systypes t on t.usertype = sc.usertype
WHERE so.name like '#myTmp%' -- table name
ORDER BY colid

-- table & column names from fields 
select so.name, sc.name, t.name, sc.length 
from syscolumns sc 
JOIN sysobjects so ON sc.id = so.id
JOIN systypes t on t.usertype = sc.usertype
where sc.name like '%ibm_logn%'
order by so.name, sc.name

-- with protection information, too
select so.name, sc.name, t.name, sc.length, su.name, sp.uid, sp.action, sp.protecttype, sp.columns
from syscolumns sc 
JOIN sysobjects so ON so.id = sc.id
JOIN systypes t on t.usertype = sc.usertype
JOIN sysprotects sp ON sp.id = sc.id
JOIN sysusers su ON su.uid = sp.uid
where sc.name like '%logn%' -- field name
AND sp.uid = 0 -- skip to see non-public tables, too
order by so.name, sc.name, su.name

