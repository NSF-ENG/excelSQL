VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Panels"
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
     .description = "Panels considering chosen projects"
     .tip = "Unhide to jump to " & .name
     .tabtitle = .description
     .note = "Panels that considered num>0 of the chosen projects.  The summary of recommendations, panlSumm, if given, reports all proposals on a panel, not just those chosen."
     .helpText = "Lists all panels considering the chosen proposals, with the number of them on each panel. If the panel has met, summarizes the recommendations for all proposals (not just the chosen)."
     .helpvideo = ""
     .phase = 1
     .order = 400
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
Dim myPanls As String
Dim mainQuery As String

'Panels with these proposals
myPanls = "SELECT panl_prop.panl_id, panl.panl_name, panl.pm_logn_id, Count(prop.lead) AS numProps, panl.panl_bgn_date, panl.panl_loc, panl.pgm_ele_code, panl.fund_org_code, panl.fund_pgm_ele_code" & vbNewLine _
& "INTO #myPanls FROM #myProps prop" & vbNewLine _
& "JOIN csd.panl_prop panl_prop ON prop.prop_id = panl_prop.prop_id " & vbNewLine _
& "JOIN csd.panl panl ON panl_prop.panl_id = panl.panl_id" & vbNewLine _
& "WHERE prop.ILN < 'M' " & vbNewLine _
& "GROUP BY panl_prop.panl_id, panl.panl_name, panl.pm_logn_id, panl.panl_bgn_date, panl.panl_loc, panl.pgm_ele_code, panl.fund_org_code, panl.fund_pgm_ele_code" & vbNewLine

mainQuery = "SELECT " & limit & " mp.*, convert(varchar,SUM(ps.rtCount)) + ' rated projects: ' +" & vbNewLine _
& "   MAX( CASE ps.RCOM_SEQ_NUM WHEN 1 THEN        convert(varchar,ps.rtCount) + ' ' +  ps.RCOM_ABBR ELSE '' END ) +" & vbNewLine _
& "   MAX( CASE ps.RCOM_SEQ_NUM WHEN 2 THEN ', ' + convert(varchar,ps.rtCount) + ' ' +  ps.RCOM_ABBR ELSE '' END ) +" & vbNewLine _
& "   MAX( CASE ps.RCOM_SEQ_NUM WHEN 3 THEN ', ' + convert(varchar,ps.rtCount) + ' ' +  ps.RCOM_ABBR ELSE '' END ) +" & vbNewLine _
& "   MAX( CASE ps.RCOM_SEQ_NUM WHEN 4 THEN ', ' + convert(varchar,ps.rtCount) + ' ' +  ps.RCOM_ABBR ELSE '' END ) +" & vbNewLine _
& "   MAX( CASE ps.RCOM_SEQ_NUM WHEN 5 THEN ', ' + convert(varchar,ps.rtCount) + ' ' +  ps.RCOM_ABBR ELSE '' END ) +" & vbNewLine _
& "   MAX( CASE ps.RCOM_SEQ_NUM WHEN 6 THEN ', ' + convert(varchar,ps.rtCount) + ' ' +  ps.RCOM_ABBR ELSE '' END ) as 'panlSumm'" & vbNewLine _
& "FROM #myPanls mp" & vbNewLine _
& "LEFT OUTER JOIN (SELECT pl.panl_id,  panl_rcom_def.RCOM_SEQ_NUM, panl_rcom_def.RCOM_ABBR, Count(panl_prop_summ.PROP_ID) AS rtCount" & vbNewLine _
& "    FROM (SELECT DISTINCT panl_id FROM #myPanls) pl" & vbNewLine _
& "    JOIN FLflpdb.flp.panl_prop_summ panl_prop_summ ON pl.panl_id = panl_prop_summ.PANL_ID" & vbNewLine _
& "    JOIN FLflpdb.flp.panl_rcom_def panl_rcom_def ON panl_prop_summ.PANL_ID = panl_rcom_def.PANL_ID AND panl_prop_summ.RCOM_SEQ_NUM = panl_rcom_def.RCOM_SEQ_NUM" & vbNewLine _
& "    JOIN csd.prop pr ON panl_prop_summ.PROP_ID = pr.prop_id AND pr.prop_id=isnull(pr.lead_prop_id,pr.prop_id) " & vbNewLine _
& "    GROUP BY pl.panl_id, panl_rcom_def.RCOM_SEQ_NUM, panl_rcom_def.RCOM_ABBR ) ps ON mp.panl_id = ps.panl_id" & vbNewLine _
& "GROUP BY mp.panl_id  ORDER BY mp.panl_bgn_date" & vbNewLine

dropTables = "DROP TABLE #myProps DROP TABLE #myPanls"

Call QT.doQuery(mInputTab.InputSQL & myPanls & mainQuery & dropTables)
End Sub
