VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ProjSumKey"
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
     .description = "Project summaries & keywords"
     .tip = "Unhide to Jump to " & .name
     .tabtitle = .description
     .note = "Retrieves text project summaries for search. Reports keywords, if present."
     .helpText = .note & " You can paste in summaries entered as pdf, then set Ref? to n for No Refresh."
     .helpvideo = ""
     .phase = 3
     .order = 1011
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
Dim myPanls, mainQuery As String

'Projects in these proposals
myPanls = "CREATE INDEX myProps_ix ON #myProps(lead)" & vbNewLine _
& "SELECT mp.lead, COUNT(panl_id) AS nPanl, MAX(panl_id) AS panl_idmax" & vbNewLine _
& "INTO #myPanlCnt FROM #myProps mp" & vbNewLine _
& "JOIN csd.panl_prop pp ON pp.prop_id = mp.lead" & vbNewLine _
& "WHERE mp.ILN < 'M' GROUP BY mp.lead" & vbNewLine & vbNewLine

mainQuery = "SELECT  " & limit & " nsf_rcvd_date, dd_rcom_date, mp.pgm_annc_id, Dir, mp.org_code,  Pgm,  PO, mp.obj_clas_code, natr_rqst_abbr, prop_stts_abbr," & vbNewLine _
& "ILN, mp.Div, mp.prop_id, pi_last_name, pi_frst_name, inst_name, st_code, pi_emai_addr, pi_id, inst_id, " & vbNewLine _
& "prop_titl_txt, (SELECT sum(p.rqst_dol) FROM #myProps p where p.lead = mp.lead) AS rqst_dol_tot, mp.rqst_eff_date, mp.rqst_mnth_cnt," & vbNewLine _
& "s.SPCL_CHAR_PDF, s.PROJ_SUMM_TXT, s.INTUL_MERT, s.BRODR_IMPT" & vbNewLine _
& "FROM #myProps mp" & vbNewLine _
& "JOIN csd.pgm_annc pgm_annc ON mp.pgm_annc_id = pgm_annc.pgm_annc_id" & vbNewLine _
& "JOIN csd.org org ON mp.org_code = org.org_code" & vbNewLine _
& "JOIN FLflpdb.flp.obj_clas_pars oc ON mp.obj_clas_code = oc.obj_clas_code" & vbNewLine _
& "JOIN FLflpdb.flp.prop_subm_ctl psc ON mp.prop_id = psc.PROP_ID" & vbNewLine _
& "LEFT OUTER JOIN FLflpdb.flp.proj_summ s ON psc.TEMP_PROP_ID = s.TEMP_PROP_ID" & vbNewLine _
& "WHERE mp.ILN < 'M'" & vbNewLine _
& "ORDER BY mp.lead" & vbNewLine

dropTables = "DROP TABLE #myProps"

Call QT.doQuery(mInputTab.InputSQL() & mainQuery & dropTables)
End Sub
