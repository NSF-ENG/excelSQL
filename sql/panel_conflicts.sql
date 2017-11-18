-- all panelists conflicts for panel minutes 

--------single table version
DECLARE @panl char(7)
SELECT @panl = ?
SELECT revr.revr_id, rtrim(revr.revr_last_name) + ', ' + rtrim(revr.revr_frst_name) AS Panelist,
isnull(c.Proposals,'-none- ') AS Proposals
FROM csd.panl_revr p
JOIN csd.revr revr ON revr.revr_id = p.revr_id
LEFT JOIN (SELECT pr.revr_id,
     MAX(CASE r.seq WHEN  1 THEN r.prop_id ELSE '' END)+
     MAX(CASE r.seq WHEN  2 THEN ', '+r.prop_id ELSE '' END)+
     MAX(CASE r.seq WHEN  3 THEN ', '+r.prop_id ELSE '' END)+
     MAX(CASE r.seq WHEN  4 THEN ', '+r.prop_id ELSE '' END)+
     MAX(CASE r.seq WHEN  5 THEN ', '+r.prop_id ELSE '' END)+
     MAX(CASE r.seq WHEN  6 THEN ', '+r.prop_id ELSE '' END)+
     MAX(CASE r.seq WHEN  7 THEN ', '+r.prop_id ELSE '' END)+
     MAX(CASE r.seq WHEN  8 THEN ', '+r.prop_id ELSE '' END)+
     MAX(CASE r.seq WHEN  9 THEN ', '+r.prop_id ELSE '' END)+
     MAX(CASE r.seq WHEN 10 THEN ', '+r.prop_id ELSE '' END)+
     MAX(CASE r.seq WHEN 11 THEN ', '+r.prop_id ELSE '' END)+
     MAX(CASE r.seq WHEN 12 THEN ', '+r.prop_id ELSE '' END)+
     MAX(CASE r.seq WHEN 13 THEN ', '+r.prop_id ELSE '' END)+
     MAX(CASE r.seq WHEN 14 THEN ', '+r.prop_id ELSE '' END)+
     MAX(CASE r.seq WHEN 15 THEN ', '+r.prop_id ELSE '' END)+
     MAX(CASE r.seq WHEN 16 THEN ', '+r.prop_id ELSE '' END)+
     MAX(CASE r.seq WHEN 17 THEN ', '+r.prop_id ELSE '' END)+
     MAX(CASE r.seq WHEN 18 THEN ', '+r.prop_id ELSE '' END)+
     MAX(CASE r.seq WHEN 19 THEN ', '+r.prop_id ELSE '' END)+
     MAX(CASE r.seq WHEN 20 THEN ', '+r.prop_id ELSE '' END) AS Proposals
    FROM csd.panl_revr pr
    JOIN (SELECT rp.revr_id, rp.prop_id, 
            (SELECT COUNT(*) 
                FROM csd.rev_prop rp2 
                JOIN csd.panl_prop pp2 ON pp2.prop_id = rp2.prop_id AND pp2.panl_id = @panl
                WHERE rp2.revr_id = rp.revr_id AND rp2.rev_stts_code = 'C' AND rp2.prop_id <= rp.prop_id
             ) AS seq 
         FROM csd.rev_prop rp 
        JOIN csd.panl_prop pp ON pp.prop_id = rp.prop_id AND pp.panl_id = @panl
        WHERE rp.rev_stts_code = 'C') r ON r.revr_id = pr.revr_id 
    GROUP BY pr.revr_id ) c ON c.revr_id = p.revr_id
WHERE p.panl_id = @panl 
ORDER BY Panelist

-- temp table version
--DROP TABLE #pConfl
DECLARE @panl char(7)
SELECT @panl = 'p180207'
SELECT pr.revr_id, pp.prop_id, id=identity(18), 0 as seq
INTO #pConfl
FROM csd.panl_revr pr 
JOIN csd.panl_prop pp ON pr.panl_id = pp.panl_id 
JOIN csd.rev_prop r ON r.revr_id = pr.revr_id AND r.prop_id = pp.prop_id AND r.rev_stts_code = 'C'
WHERE pr.panl_id = ?
ORDER BY pr.revr_id, pp.prop_id

SELECT revr_id,MIN(id) as start INTO #myStarts FROM #pConfl GROUP BY revr_id
UPDATE #pConfl SET seq = id-M.start 
    FROM #pConfl r, #myStarts M  WHERE M.revr_id = r.revr_id
DROP TABLE #myStarts


SELECT pr.revr_id,  rtrim(revr.revr_last_name) + ', ' + rtrim(revr.revr_frst_name) AS Panelist,
isnull(c.Proposals,'-none- ') as Proposals
FROM csd.panl_revr pr
JOIN csd.revr revr ON revr.revr_id = pr.revr_id
LEFT JOIN (SELECT revr_id,
     MAX(CASE seq WHEN  0 THEN prop_id ELSE '' END)+
     MAX(CASE seq WHEN  1 THEN ', '+prop_id ELSE '' END)+
     MAX(CASE seq WHEN  2 THEN ', '+prop_id ELSE '' END)+
     MAX(CASE seq WHEN  3 THEN ', '+prop_id ELSE '' END)+
     MAX(CASE seq WHEN  4 THEN ', '+prop_id ELSE '' END)+
     MAX(CASE seq WHEN  5 THEN ', '+prop_id ELSE '' END)+
     MAX(CASE seq WHEN  6 THEN ', '+prop_id ELSE '' END)+
     MAX(CASE seq WHEN  7 THEN ', '+prop_id ELSE '' END)+
     MAX(CASE seq WHEN  8 THEN ', '+prop_id ELSE '' END)+
     MAX(CASE seq WHEN  9 THEN ', '+prop_id ELSE '' END)+
     MAX(CASE seq WHEN 10 THEN ', '+prop_id ELSE '' END)+
     MAX(CASE seq WHEN 11 THEN ', '+prop_id ELSE '' END)+
     MAX(CASE seq WHEN 12 THEN ', '+prop_id ELSE '' END)+
     MAX(CASE seq WHEN 13 THEN ', '+prop_id ELSE '' END)+
     MAX(CASE seq WHEN 14 THEN ', '+prop_id ELSE '' END)+
     MAX(CASE seq WHEN 15 THEN ', '+prop_id ELSE '' END)+
     MAX(CASE seq WHEN 16 THEN ', '+prop_id ELSE '' END)+
     MAX(CASE seq WHEN 17 THEN ', '+prop_id ELSE '' END)+
     MAX(CASE seq WHEN 18 THEN ', '+prop_id ELSE '' END)+
     MAX(CASE seq WHEN 19 THEN ', '+prop_id ELSE '' END)+
     MAX(CASE seq WHEN 20 THEN ', '+prop_id ELSE '' END) AS Proposals 
    FROM #pConfl 
    GROUP BY revr_id) c ON c.revr_id = pr.revr_id
WHERE pr.panl_id = ?
ORDER BY Panelist

DROP TABLE #pConfl
