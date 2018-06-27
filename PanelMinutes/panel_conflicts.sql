-- all panelists Cfl for panel minutes 

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
WHERE pr.panl_id = @panl
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
WHERE pr.panl_id = @panl
ORDER BY Panelist

DROP TABLE #pConfl
-----
-- multi-panel temp table version
--[pm_panls
SELECT pr.panl_id, pr.revr_id, pp.prop_id, id=identity(18), 0 as seq
INTO #pConfl
FROM csd.panl_revr pr 
JOIN csd.panl_prop pp ON pr.panl_id = pp.panl_id 
JOIN csd.rev_prop r ON r.revr_id = pr.revr_id AND r.prop_id = pp.prop_id AND r.rev_stts_code = 'C'
WHERE pr.panl_id IN
--]pm_panls
('p181992','p181938','p181917','P160845','P160844','P160854')
--[pm_panls2

ORDER BY pr.revr_id, pp.prop_id

SELECT revr_id,MIN(id) as start INTO #myStarts FROM #pConfl GROUP BY revr_id
UPDATE #pConfl SET seq = id-M.start 
    FROM #pConfl r, #myStarts M  WHERE M.revr_id = r.revr_id
DROP TABLE #myStarts

--select * from  #pConfl

SELECT cf.panl_id, cf.revr_id, rtrim(revr.revr_last_name) + ', ' + rtrim(revr.revr_frst_name) + ': ' + cf.Cfl as Cfl, id=identity(18), 0 as seq
INTO #myConfl
FROM(SELECT c.panl_id, c.revr_id, MAX(CASE seq WHEN  0 THEN prop_id ELSE '' END)+
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
     MAX(CASE seq WHEN 20 THEN ', '+prop_id ELSE '' END)+char(10) AS Cfl
    FROM #pConfl c
    GROUP BY c.panl_id, c.revr_id) cf
JOIN csd.revr revr ON revr.revr_id = cf.revr_id
ORDER BY cf.panl_id, Cfl

SELECT panl_id,MIN(id) as start INTO #mySt2 FROM #myConfl GROUP BY panl_id
UPDATE #myConfl SET seq = id-M.start 
    FROM #myConfl r, #mySt2 M  WHERE M.panl_id = r.panl_id
DROP TABLE #mySt2

--select * from #myConfl
--drop table  #myConfl
--drop table 


SELECT panl.panl_id, panl.panl_name, panl.panl_bgn_date, panl.panl_end_date, panl.panl_loc, panl.org_code, panl.pgm_ele_code, panl.pm_logn_id,panl.meet_type_code, panl.meet_fmt, panl.fund_org_code, panl.fund_pgm_ele_code, panl.fund_app_code, panl_stts.panl_stts_txt, panl.oblg_flag, org.org_long_name, pgm_ele.pgm_ele_long_name,
(SELECT COUNT(panl_prop.prop_id) FROM csd.panl_prop panl_prop, csd.prop prop WHERE panl.panl_id = panl_prop.panl_id AND panl_prop.prop_id = prop.prop_id AND (prop.prop_id = isnull( prop.lead_prop_id, prop.prop_id) ) ) AS 'Nproj',
(SELECT COUNT(panl_prop.prop_id) FROM csd.panl_prop panl_prop WHERE panl.panl_id = panl_prop.panl_id) AS 'Nprop',
nullif((SELECT COUNT(panl_revr.revr_id) FROM csd.panl_revr panl_revr WHERE panl.panl_id = panl_revr.panl_id AND (panl_revr.tele_conf_part_flag = 'N')),0) AS 'Nrevr',
nullif((SELECT COUNT(panl_revr.revr_id) FROM csd.panl_revr panl_revr WHERE panl.panl_id = panl_revr.panl_id AND (panl_revr.tele_conf_part_flag = 'Y')),0) AS 'Nvirt_revr',
    (SELECT MAX(CASE seq WHEN  0 THEN Cfl ELSE '' END)+
     MAX(CASE seq WHEN  1 THEN Cfl ELSE '' END)+
     MAX(CASE seq WHEN  2 THEN Cfl ELSE '' END)+
     MAX(CASE seq WHEN  3 THEN Cfl ELSE '' END)+
     MAX(CASE seq WHEN  4 THEN Cfl ELSE '' END)+
     MAX(CASE seq WHEN  5 THEN Cfl ELSE '' END)+
     MAX(CASE seq WHEN  6 THEN Cfl ELSE '' END)+
     MAX(CASE seq WHEN  7 THEN Cfl ELSE '' END)+
     MAX(CASE seq WHEN  8 THEN Cfl ELSE '' END)+
     MAX(CASE seq WHEN  9 THEN Cfl ELSE '' END)+
     MAX(CASE seq WHEN 10 THEN Cfl ELSE '' END)+
     MAX(CASE seq WHEN 11 THEN Cfl ELSE '' END)+
     MAX(CASE seq WHEN 12 THEN Cfl ELSE '' END)+
     MAX(CASE seq WHEN 13 THEN Cfl ELSE '' END)+
     MAX(CASE seq WHEN 14 THEN Cfl ELSE '' END)+
     MAX(CASE seq WHEN 15 THEN Cfl ELSE '' END)+
     MAX(CASE seq WHEN 16 THEN Cfl ELSE '' END)+
     MAX(CASE seq WHEN 17 THEN Cfl ELSE '' END)+
     MAX(CASE seq WHEN 18 THEN Cfl ELSE '' END)+
     MAX(CASE seq WHEN 19 THEN Cfl ELSE '' END)+
     MAX(CASE seq WHEN 20 THEN Cfl ELSE '' END)+
     MAX(CASE seq WHEN 21 THEN Cfl ELSE '' END)+
     MAX(CASE seq WHEN 22 THEN Cfl ELSE '' END)+
     MAX(CASE seq WHEN 23 THEN Cfl ELSE '' END)+
     MAX(CASE seq WHEN 24 THEN Cfl ELSE '' END)+
     MAX(CASE seq WHEN 25 THEN Cfl ELSE '' END)+
     MAX(CASE seq WHEN 26 THEN Cfl ELSE '' END)+
     MAX(CASE seq WHEN 27 THEN Cfl ELSE '' END)+
     MAX(CASE seq WHEN 28 THEN Cfl ELSE '' END)+
     MAX(CASE seq WHEN 29 THEN Cfl ELSE '' END)+
     MAX(CASE seq WHEN 30 THEN Cfl ELSE '' END)+
     MAX(CASE seq WHEN 31 THEN Cfl ELSE '' END)+
     MAX(CASE seq WHEN 32 THEN Cfl ELSE '' END)+
     MAX(CASE seq WHEN 33 THEN Cfl ELSE '' END)+
     MAX(CASE seq WHEN 34 THEN Cfl ELSE '' END)+
     MAX(CASE seq WHEN 35 THEN Cfl ELSE '' END)+
     MAX(CASE seq WHEN 36 THEN Cfl ELSE '' END)+
     MAX(CASE seq WHEN 37 THEN Cfl ELSE '' END)+
     MAX(CASE seq WHEN 38 THEN Cfl ELSE '' END)+
     MAX(CASE seq WHEN 39 THEN Cfl ELSE '' END)+
     MAX(CASE seq WHEN 40 THEN Cfl ELSE '' END) 
FROM #myConfl c WHERE c.panl_id = panl.panl_id) as Conflicts
FROM csd.org org, csd.panl panl, csd.panl_stts panl_stts, csd.pgm_ele pgm_ele
WHERE panl.panl_stts_code = panl_stts.panl_stts_code AND panl.pgm_ele_code = pgm_ele.pgm_ele_code AND panl.org_code = org.org_code AND 
panl.panl_id In 
--]pm_panls2
('p181992','p181938','p181917','P160845','P160844','P160854')
--[pm_drop
DROP TABLE #pConfl,#myConfl
--]pm_drop
