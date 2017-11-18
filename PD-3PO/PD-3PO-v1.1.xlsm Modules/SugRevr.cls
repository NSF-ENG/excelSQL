VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SugRevr"
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
     .description = "Proposals with suggested or unwanted reviewers"
     .tip = "Unhide to jump to " & .name
     .tabtitle = "Suggested & unwanted reviewers"
     .note = "Collects PIs' suggested and unwanted reviewers for those proposals that have indicated them."
     .helpText = "Reports the suggested and unwanted reviewers that PIs have listed in the correponding single copy documents.  Sometimes the reasons make for interesting reading."
     .helpvideo = ""
     .phase = 3
     .order = 3003
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

'Suggested Reviewers and Unwanted Reviewers these proposals
mainQuery = "SELECT prop.nsf_rcvd_date, Dir, org_code, Pgm, PO, prop.natr_rqst_abbr as natr_rqst, prop.prop_stts_abbr as prop_stts," & vbNewLine _
& "prop.ILN, prop.lead, Div, prop.prop_id, prop.pi_last_name, prop.pi_frst_name, prop.inst_name," & vbNewLine _
& "srevr.revr_want as Suggested_Reviewers, srevr.revr_dont_want as Unwanted_Reviewers, " & vbNewLine _
& "prop.prop_titl_txt, len(srevr.revr_want)+len(srevr.revr_dont_want) as len_tot" & vbNewLine _
& "FROM #myProps prop " & vbNewLine _
& "JOIN csd.prop_subm_ctl_vw ctl ON prop.PROP_ID=ctl.PROP_ID " & vbNewLine _
& "JOIN csd.sugg_revr_vw srevr ON ctl.TEMP_PROP_ID=srevr.TEMP_PROP_ID " & vbNewLine _
& "ORDER BY prop.lead, prop.ILN, prop.prop_id" & vbNewLine
   
'MsgBox myReviewers
dropTables = "DROP TABLE #myProps"

Call QT.doQuery(mInputTab.InputSQL & mainQuery & dropTables)

End Sub
