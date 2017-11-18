VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Orphans"
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
     .description = "Report 'Orphans': proposals not on panels"
     .tip = "Check prop_stts"
     .tabtitle = "Orphan Check: proposals not paneled"
     .note = "Proposals not paneled; possible reasons in cols F-K: ad hoc only, object class is not Research Grant, request is SUPP,TRSF, status is WTH,RTNR"
     .helpText = "Lists all proposals that are not paneled.  Possible reasons may be found in cols F-K:  1) reviewed by ad hocs, 2) object class is conference, etc, 3) request is Supplement, Transfer, etc. 4) status is withdrawn or returned."
     .helpvideo = ""
     .phase = 3
     .order = 1101
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
Dim mainQuery As String

'Projects in these proposals
mainQuery = "SELECT nsf_rcvd_date, Dir, org_code, Pgm, PO,  nullif(Count(rev_prop.rev_sent_date),0) as sent_adhoc, nullif(Count(rev_prop.rev_sent_date)-Count(rev_prop.rev_rtrn_date),0) as out_adhoc, MAX(rev_prop.rev_due_date) as last_due," & vbNewLine _
& "obj_clas_code, natr_rqst_abbr, prop_stts_abbr, ILN, lead, Div, prop.prop_id, pi_last_name, pi_frst_name, inst_name, prop_titl_txt, prop_stts_txt" & vbNewLine _
& "FROM #myProps prop" & vbNewLine _
& "LEFT OUTER JOIN  csd.rev_prop rev_prop ON prop.prop_id = rev_prop.prop_id AND rev_prop.rev_type_code='R'" & vbNewLine _
& "WHERE not exists (select pp.panl_id FROM csd.panl_prop pp WHERE prop.prop_id = pp.prop_id)" & vbNewLine _
& "GROUP BY prop_stts_txt, Pgm, PO, natr_rqst_abbr, Dir, ILN, lead, Div, prop.prop_id, pi_last_name, pi_frst_name, inst_name, prop_titl_txt, nsf_rcvd_date, prop_stts_abbr" & vbNewLine _
& "ORDER BY prop_stts_abbr, lead, ILN" & vbNewLine

dropTables = "DROP TABLE #myProps"

Call QT.doQuery(mInputTab.InputSQL() & mainQuery & dropTables)
End Sub


