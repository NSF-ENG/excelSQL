VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "NewInst"
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
     .description = "[slow]Pend&Rcom from New Awardee Institutions"
     .tip = "A little slow; often none returned."
     .tabtitle = "New Awardee Institutions"
     .note = "For the chosen proposals that are pending, checks if any institutions have not held an NSF grant in the last 58 months. These may have additional paperwork before an award.  [Resource intensive; keep hidden]"
     .helpText = "Returns the pending and recommended proposals those for which the lead institution has not managed a funded award in the last 58 months.  New awardee institutions (none in last 60 months) have extra paperwork to be able to receive an award."
     .helpvideo = ""
     .phase = 3
     .order = 1201
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
Dim Query As String

' Find institiutions submitting pending & recommended proposals that have not held an NSF grant in previous 58 months.
' Report their proposals.  No reason to limit this query

Query = "SELECT p.nsf_rcvd_date,  p.Dir, p.org_code, Pgm, PO, newInst.inst_id, inst.inst_name, inst.st_code, inst.ctry_code,  " & vbNewLine _
& "inst.awd_perf_inst_code, inst.duns_id, inst.perf_org_code, perf_org.perf_org_txt as perf_org_type, " & vbNewLine _
& "prop_stts_abbr, natr_rqst_abbr, Div, p.prop_id, pi_last_name, pi_frst_name, inst.inst_shrt_name,  " & vbNewLine _
& "prop_stts_txt, p.prop_titl_txt" & vbNewLine _
& "FROM (SELECT pendInst.inst_id" & vbNewLine _
& "    FROM (SELECT distinct prop.inst_id FROM #myProps prop WHERE prop.prop_stts_abbr in ('PEND','RCOM') ) pendInst" & vbNewLine _
& "    WHERE NOT EXISTS (SELECT * FROM csd.awd awd WHERE awd.AWD_EXP_DATE >= dateadd(mm, -58, getdate()) AND awd.INST_ID = pendInst.INST_ID ) ) newInst" & vbNewLine _
& "LEFT OUTER JOIN csd.inst inst ON inst.inst_id = newInst.inst_id" & vbNewLine _
& "LEFT OUTER JOIN csd.perf_org perf_org ON perf_org.perf_org_code = inst.perf_org_code " & vbNewLine _
& "JOIN #myProps p ON p.inst_id = newInst.inst_id " & vbNewLine _
& "ORDER BY inst_name" & vbNewLine

dropTables = "DROP TABLE #myProps" & vbNewLine

Call QT.doQuery(mInputTab.InputSQL() & Query & dropTables)
End Sub


