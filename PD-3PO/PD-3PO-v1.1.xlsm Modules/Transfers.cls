VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Transfers"
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
     .description = "[slow]All transfers for proposals"
     .tip = "Unhide to jump to " & .name
     .tabtitle = "Proposal Transfer History"
     .note = "For the chosen proposals and thier related awards, list transfers and changes between programs, PDs, institutions, and PIs. Can be slow."
     .helpText = "Reports all proposal and award transfers for the listed proposals. Note: this is proposal-based; for award management see the award-based query."
     .helpvideo = ""
     .phase = 3
     .order = 1451
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
Dim getAwds, getXfers, Query As String

'all related transfers throught the life cycle of the involved proposals
getAwds = "create index myProps_ix ON #myProps(prop_id)" & vbNewLine _
& "select p.Pgm, p.natr_rqst_abbr, p.prop_stts_abbr, p.ILN, p.lead,p.Div, p.prop_id, a.awd_id" & vbNewLine _
& "into #myPropsAwds" & vbNewLine _
& "from #myProps p  " & vbNewLine _
& "JOIN csd.amd a ON p.prop_id = a.prop_id " & vbNewLine _
& "UNION select p.Pgm, p.natr_rqst_abbr, p.prop_stts_abbr, p.ILN, p.lead, p.Div, p.prop_id, c.PREV_AWD_ID " & vbNewLine _
& "from #myProps p " & vbNewLine _
& "join csd.prop_subm_ctl_vw ctl on p.prop_id = ctl.prop_id " & vbNewLine _
& "join csd.prop_covr_vw c on c.TEMP_PROP_ID = ctl.TEMP_PROP_ID " & vbNewLine _
& "WHERE c.PREV_AWD_ID <> '' AND p.natr_rqst_abbr not in ('NEW','RNEW','ABR') " & vbNewLine '-- skip PI supplied prev awad info

'-- get award transfers, which give other relevant proposals
getXfers = "SELECT distinct a.Pgm, a.natr_rqst_abbr, a.prop_stts_abbr, a.ILN, a.lead, a.Div, a.prop_id, a.awd_id, min(l.last_prop_id) as last_prop_id, pi.pi_last_name, inst.inst_shrt_name, 'awd ' as xfer_type, " & vbNewLine _
& "    CASE WHEN l.org_code <> prop.org_code THEN prop.org_code END as fromORG, " & vbNewLine _
& "    CASE WHEN l.pgm_ele_code <> prop.pgm_ele_code THEN prop.pgm_ele_code END as fromPEC," & vbNewLine _
& "    CASE WHEN l.pm_ibm_logn_id <> prop.pm_ibm_logn_id THEN prop.pm_ibm_logn_id END as fromPA_PO,  " & vbNewLine _
& "    l.org_code as toORG, l.pgm_ele_code as toPEC, l.pm_ibm_logn_id as toPA_PO, min(l.awd_chg_date) as crtd_date, convert(varchar(255),null) as note" & vbNewLine _
& "INTO #myXfers" & vbNewLine _
& "FROM #myPropsAwds a " & vbNewLine _
& "JOIN csd.awd_chg_log l ON l.awd_id = a.awd_id" & vbNewLine _
& "JOIN csd.prop prop ON prop.prop_id = l.last_prop_id" & vbNewLine _
& "JOIN csd.pi_vw pi ON prop.pi_id = pi.pi_id" & vbNewLine _
& "JOIN csd.inst inst ON prop.inst_id = inst.inst_id" & vbNewLine _
& "WHERE a.awd_id <> '' " & vbNewLine _
& "Group by a.Pgm, a.prop_stts_abbr, a.natr_rqst_abbr, a.lead, a.ILN, a.Div, a.prop_id, a.awd_id, pi.pi_last_name, inst.inst_shrt_name, l.pm_ibm_logn_id, l.org_code, l.pgm_ele_code, prop.pm_ibm_logn_id, prop.org_code, prop.pgm_ele_code" & vbNewLine

'-- get all proposal transfers;  both those involved in Xfers and those originally pulled
getXfers = getXfers & "insert into #myXfers SELECT x.*, 'prop' as xfer_type, " & vbNewLine _
& "    CASE WHEN to_org_code <> from_org_code THEN from_org_code END as fromORG, " & vbNewLine _
& "    CASE WHEN to_pgm_ele_code <> from_pgm_ele_code THEN from_pgm_ele_code END as fromPEC," & vbNewLine _
& "    CASE WHEN to_lan_id <> from_lan_id THEN from_lan_id END as fromPA_PO, " & vbNewLine _
& "    to_org_code as toORG, to_pgm_ele_code as toPEC, to_lan_id as toPA_PO, h.crtd_date, isnull(h.note,'') as note " & vbNewLine _
& "FROM ( SELECT a.Pgm, a.natr_rqst_abbr, a.prop_stts_abbr,  a.ILN,a.lead, a.Div, a.prop_id, a.awd_id, a.last_prop_id, a.pi_last_name, a.inst_shrt_name" & vbNewLine _
& "       FROM #myXfers a " & vbNewLine _
& "       WHERE a.prop_id <> a.last_prop_id " & vbNewLine _
& " UNION SELECT p.Pgm, p.natr_rqst_abbr, p.prop_stts_abbr,  p.ILN, p.lead, p.Div, p.prop_id, a.awd_id, p.prop_id, p.pi_last_name, p.inst_name " & vbNewLine _
& "       FROM #myProps p " & vbNewLine _
& "       LEFT OUTER JOIN #myPropsAwds a ON p.prop_id = a.prop_id ) x" & vbNewLine _
& "JOIN FLflpdb.flp.prop_ownr_hist h ON h.prop_id = x.last_prop_id " & vbNewLine _
& "WHERE h.ownr_stts_code = 9 " & vbNewLine

Query = "select " & limit & " * from #myXfers order by lead, ILN, prop_id, awd_id, crtd_date" & vbNewLine

dropTables = "DROP TABLE #myProps DROP TABLE #myPropsAwds DROP TABLE #myXfers" & vbNewLine

Call QT.doQuery(mInputTab.InputSQL & getAwds & getXfers & Query & dropTables)
End Sub
