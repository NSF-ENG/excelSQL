VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Proposals"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit
Private QT As clsQueryTable

Function Descriptor() As clsSheetDescriptor
' return descriptor for this sheet
    Set Descriptor = New clsSheetDescriptor
    With Descriptor
     .name = Me.name
     .description = "Proposal-specific (EPSCoR, Hum/Vrt, PRC, route PECs...)"
     .tip = "Unhide to Jump to " & .name
     .tabtitle = "Proposal Specific Information"
     .note = "Requested amount, coversheet items,  (human/vertibrates & international), routing PECs, PRCs. Note: most good stuff is on the Project or PersInst tabs."
     .helpText = "For projects or personnel use other tabs, but this pulls proposal specific details like requested amounts, coversheet items (human/vertibrates & international), routing PECs, and PRCs."
     .helpvideo = ""
     .phase = 1
     .order = 300
     .uigroup = .order Mod 10
     .depends = InputTab.name
    End With
End Function

Sub InitializeQueryTable()
' record query table for use by QT.doRefresh and events BeforeRefresh and AfterRefresh
' Called in ThisWorkbook_Open with on Error resume next so turn that off first.
  On Error GoTo 0
  Set QT = New clsQueryTable
  Call QT.initialize(Me, Me.ListObjects(1).QueryTable)
End Sub

Sub ClearRS()
  If Not QT Is Nothing Then Call QT.clear
End Sub

Sub RefreshRS(Optional limit As String = "")
'refresh tab based on parameters in InputTab
Dim dropTables As String
Dim myCtrys, myPRCs, mainQuery As String

'allPIs on these proposals

myCtrys = "SELECT sp1.TEMP_PROP_ID, ctry.ctry_name, id=identity(18), 0 as 'seq' INTO #myCtry" & vbNewLine _
& "FROM #myProps prop" & vbNewLine _
& "JOIN csd.prop_subm_ctl_vw psc ON prop.prop_id = psc.prop_id" & vbNewLine _
& "JOIN csd.prop_spcl_item_vw sp1 ON sp1.TEMP_PROP_ID = psc.TEMP_PROP_ID" & vbNewLine _
& "JOIN csd.ctry ctry ON sp1.SPCL_ITEM_CODE = ctry.ctry_code" & vbNewLine _
& "WHERE end_date Is Null" & vbNewLine _
& "ORDER BY sp1.TEMP_PROP_ID, ctry.ctry_name" & vbNewLine _
& "SELECT TEMP_PROP_ID, MIN(id) as 'start' INTO #myStart FROM #myCtry GROUP BY TEMP_PROP_ID" & vbNewLine _
& "UPDATE #myCtry set seq = id-M.start FROM #myCtry r, #myStart M WHERE r.TEMP_PROP_ID = M.TEMP_PROP_ID" & vbNewLine _
& "DROP TABLE #myStart" & vbNewLine

myPRCs = "SELECT DISTINCT prop.prop_id, pa.prop_atr_code, id=identity(18), 0 as 'seq' INTO #myPRCs" & vbNewLine _
& "FROM #myProps prop, csd.prop_atr pa WHERE pa.prop_id = prop.prop_id  AND pa.prop_atr_type_code = 'PRC'" & vbNewLine _
& "ORDER BY prop.prop_id, pa.prop_atr_code" & vbNewLine _
& "SELECT prop_id, MIN(id) as 'start' INTO #mySt2 FROM #myPRCs GROUP BY prop_id" & vbNewLine _
& "UPDATE #myPRCs set seq = id-M.start FROM #myPRCs r, #mySt2 M WHERE r.prop_id = M.prop_id" & vbNewLine _
& "DROP TABLE #mySt2" & vbNewLine

mainQuery = "SELECT " & limit & " nsf_rcvd_date, dd_rcom_date, prop.pgm_annc_id, Dir, prop.org_code, Pgm, PO," & vbNewLine _
& "obj_clas_code, prop.natr_rqst_abbr, prop.prop_stts_abbr, ILN, prop.lead, Div, prop.prop_id, pi_last_name, pi_frst_name, inst_name," & vbNewLine _
& "prop.prop_titl_txt, prop.rqst_dol, prop.rqst_eff_date, prop.rqst_mnth_cnt, prop.cntx_stmt_id, prop.inst_id, prop.pi_id, st_code," & vbNewLine _
& "(SELECT MAX( CASE pa.seq WHEN 0 THEN     rtrim(pa.prop_atr_code) END)+" & vbNewLine _
& "        MAX( CASE pa.seq WHEN 1 THEN ','+rtrim(pa.prop_atr_code) END)+" & vbNewLine _
& "        MAX( CASE pa.seq WHEN 2 THEN ','+rtrim(pa.prop_atr_code) END)+ " & vbNewLine _
& "        MAX( CASE pa.seq WHEN 3 THEN ','+rtrim(pa.prop_atr_code) END)+" & vbNewLine _
& "        MAX( CASE pa.seq WHEN 4 THEN ','+rtrim(pa.prop_atr_code) END)+" & vbNewLine _
& "        MAX( CASE pa.seq WHEN 5 THEN ','+rtrim(pa.prop_atr_code) END)+" & vbNewLine _
& "        MAX( CASE pa.seq WHEN 6 THEN ','+rtrim(pa.prop_atr_code) END)" & vbNewLine _
& "        FROM #myPRCs pa WHERE pa.prop_id = prop.prop_id) AS 'PRCs'," & vbNewLine _
& "CASE WHEN PROP_COVR.HUM_DATE is not NULL THEN convert(varchar(10),PROP_COVR.HUM_DATE,1) WHEN PROP_COVR.humn_date_pend_flag='Y' THEN 'Pend' END AS humn_date," & vbNewLine _
& "CASE WHEN PROP_COVR.VERT_DATE is not NULL THEN convert(varchar(10),PROP_COVR.VERT_DATE,1) WHEN PROP_COVR.vrtb_date_pend_flag='Y' THEN 'Pend' END AS vrtb_date," & vbNewLine


mainQuery = mainQuery & "(SELECT  MAX( CASE mySeq.seq WHEN 0 THEN mySeq.CODE ELSE '' END ) + ' ' +" & vbNewLine _
& "                 MAX( CASE mySeq.seq WHEN 1 THEN mySeq.CODE ELSE '' END ) + ' ' +" & vbNewLine _
& "                MAX( CASE mySeq.seq WHEN 2 THEN mySeq.CODE ELSE '' END ) + ' ' +" & vbNewLine _
& "                MAX( CASE mySeq.seq WHEN 3 THEN mySeq.CODE ELSE '' END )" & vbNewLine _
& "  FROM (SELECT  r1.TEMP_PROP_ID, r1.CODE," & vbNewLine _
& "                            (SELECT  count(*) FROM FLflpdb.flp.routing r2" & vbNewLine _
& "                              WHERE r2.TEMP_PROP_ID = r1.TEMP_PROP_ID" & vbNewLine _
& "                                             AND r2.SEQUENCE < r1.SEQUENCE ) as 'seq'" & vbNewLine _
& "                 FROM FLflpdb.flp.routing r1) mySeq" & vbNewLine _
& "  WHERE psc.TEMP_PROP_ID = mySeq.TEMP_PROP_ID)  AS 'rout_PECs'," & vbNewLine

mainQuery = mainQuery & "(SELECT MAX(CASE b.seq WHEN 1 THEN b.ctry_name ELSE '' END)+" & vbNewLine _
& "    MAX(CASE b.seq WHEN 2 THEN '; '+b.ctry_name ELSE '' END)+" & vbNewLine _
& "    MAX(CASE b.seq WHEN 3 THEN '; '+b.ctry_name ELSE '' END)+" & vbNewLine _
& "    MAX(CASE b.seq WHEN 4 THEN '; '+b.ctry_name ELSE '' END)+" & vbNewLine _
& "    MAX(CASE b.seq WHEN 5 THEN '; '+b.ctry_name ELSE '' END)+" & vbNewLine _
& "    MAX(CASE b.seq WHEN 6 THEN '; '+b.ctry_name ELSE '' END)+" & vbNewLine _
& "    MAX(CASE b.seq WHEN 7 THEN '; '+b.ctry_name ELSE '' END)+" & vbNewLine _
& "    MAX(CASE b.seq WHEN 8 THEN '; '+b.ctry_name ELSE '' END) " & vbNewLine _
& "    FROM #myCtry b WHERE b.TEMP_PROP_ID = psc.TEMP_PROP_ID) AS Country" & vbNewLine _
& "FROM #myProps prop" & vbNewLine _
& "JOIN csd.prop_subm_ctl_vw psc ON prop.prop_id = psc.prop_id " & vbNewLine _
& "JOIN FLflpdb.flp.PROP_COVR PROP_COVR ON psc.TEMP_PROP_ID = PROP_COVR.TEMP_PROP_ID" & vbNewLine _
& "ORDER BY prop.lead, ILN, prop.prop_id" & vbNewLine & vbNewLine

dropTables = "DROP TABLE #myPRCs  DROP TABLE #myProps"

Call QT.doQuery(mInputTab.InputSQL() & myCtrys & myPRCs & mainQuery & dropTables)
End Sub

