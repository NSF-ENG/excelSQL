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
WHERE so.name like 'obj_clas%' and sc.name like '%' -- table name, field name
AND (su.name is null OR su.name = 'public'OR su.name like 'iipuser')
order by so.name, colid, su.name

SELECT so.name as tablename, so.crdate, su.name as username, sp.uid, sp.action, sp.protecttype, 
inttohex(sp.columns) as col
from sysobjects so 
LEFT JOIN sysprotects sp ON sp.id = so.id
LEFT JOIN sysusers su ON su.uid = sp.uid
WHERE so.name like '%oblg%'  -- table name, field name
--AND (su.name like 'public')
order by so.name, su.name
