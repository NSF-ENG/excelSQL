Attribute VB_Name = "mQueries"
Option Explicit

Sub BasicQueries(myPids As String)
ufProgress.Show vbModeless
Call UpdateProgressBar(0.01)
With HiddenSettings
    'need this first to get revtable in order
    Call doQuery(PRCs.ListObjects("PECGlossaryQTable").QueryTable, myPids _
        & .Range("RA_leads") & .Range("RA_PECglossary") & .Range("revtable") & "DROP TABLE #myPidRAt, #myProp, #myLead, #myRA")
Call UpdateProgressBar(0.1)
    Call doQuery(PRCs.ListObjects("PRCGlossaryQTable").QueryTable, myPids _
        & .Range("RA_leads") & .Range("RA_propPRCs") & .Range("RA_PRCglossary") _
        & "DROP TABLE #myPidRAt, #myProp, #myLead, #myRA DROP TABLE #myPRCs, #myPRCdata")
Call UpdateProgressBar(0.2)
    Call doQuery(Conflict.ListObjects("ConflictQTable").QueryTable, myPids _
        & .Range("RA_leads") & .Range("RA_ckConfRevrInst") & "DROP TABLE #myPidRAt, #myProp, #myLead, #myRA")
Call UpdateProgressBar(0.3)
    ' this is the slowest; do it in the background
    Call doQuery(ProjText.ListObjects("ProjTextQTable").QueryTable, myPids _
        & .Range("RA_leads") & .Range("RA_projText") _
        & "DROP TABLE #myPidRAt, #myProp, #myLead, #myRA DROP TABLE #myRevInfo, #mySumm")
Call UpdateProgressBar(0.5)
    Call doQuery(ckCoding.ListObjects("ckCodingQTable").QueryTable, myPids _
        & .Range("RA_leads") & .Range("RA_propPRCs") & .Range("RA_revs") & .Range("RA_budg") _
        & .Range("RA_prop") & .Range("RA_confl") & .Range("RA_panl") & .Range("RA_propCheck") & vbNewLine _
        & "DROP TABLE #myPidRAt, #myProp, #myLead, #myRA, #myPRCs, #myPRCdata DROP TABLE #myRevs, #myRevPanl, #myRevMarks, #myRevSumm " & vbNewLine _
        & "DROP TABLE #myBudg, #myPropBudg, #myPropInfo DROP TABLE #myPPConfl, #myPanl, #myProjPanl, #myProjPanlSumm")
   Call ckCodingCF
Call UpdateProgressBar(0.6)
    Call doQuery(RAData.ListObjects("RADataQTable").QueryTable, myPids _
        & .Range("RA_leads") & .Range("RA_propPRCs") & .Range("RA_revs") & .Range("RA_budg") _
        & .Range("RA_prop") & .Range("RA_confl") & .Range("RA_panl") & .Range("RA_allRAdata") & .Range("RA_allRAdata2") & vbNewLine _
        & "DROP TABLE #myPidRAt, #myProp, #myLead, #myRA, #myPRCs, #myPRCdata, #myBudg, #myPropBudg, #myPropInfo, #myRevs, #myRevPanl, #myRevMarks, #myRevSumm " & vbNewLine _
        & "DROP TABLE #myPPConfl, #myPanl, #myProjPanl, #myProjPanlSumm DROP TABLE #myDmog")
End With
Call CleanUpSheet(ckCoding)
Call CleanUpSheet(Conflict)
Call CleanUpSheet(ProjText)
Call CleanUpSheet(RAData)
End Sub

Sub AwdCodingQueries(myPids As String)
With HiddenSettings
    'these can be done for awards only
    Call doQuery(Budgets.ListObjects("BudgetsQTable").QueryTable, myPids _
        & .Range("RA_leads") & .Range("RA_propPRCs") & .Range("RA_budg") & .Range("RA_budgBlocks") _
        & "DROP TABLE #myPidRAt, #myProp, #myLead, #myRA, #myPRCs, #myPRCdata DROP TABLE #myBudg")
Call UpdateProgressBar(0.7)
        
    Call doQuery(ckAwd.ListObjects("ckAwdQTable").QueryTable, myPids _
        & .Range("RA_leads") & .Range("RA_propPRCs") & .Range("RA_budg") & .Range("RA_prop") _
        & .Range("RA_awdCheck") _
        & "DROP TABLE #myPidRAt, #myProp, #myLead, #myRA, #myPRCs, #myPRCdata DROP TABLE #myBudg, #myPropBudg, #myPropInfo DROP TABLE #myCtry, #myCovrInfo")
    Call ckAwdCF
Call UpdateProgressBar(0.8)
    
    Call doQuery(ckSplits.ListObjects("ckSplitsQTable").QueryTable, myPids _
        & .Range("RA_leads") & .Range("RA_propPRCs") & .Range("RA_budg") & .Range("RA_prop") & .Range("RA_splits") _
        & "DROP TABLE #myPidRAt, #myProp, #myLead, #myRA, #myPRCs, #myPRCdata DROP TABLE #myBudg, #myPropBudg, #myPropInfo DROP TABLE #myBSprc")
    Call ckSplitsCF
Call UpdateProgressBar(0.9)
End With
Call RefreshPivotTables(ckSplits)
Call UpdateProgressBar(1#)
Call CleanUpSheet(Budgets)
Call CleanUpSheet(ckSplits)
Call CleanUpSheet(ckAwd)
Unload ufProgress
End Sub

' Empty tables lose conditional formatting that references rows above, so clear and re-apply conditional formats for those tables.
' Note: this precludes user-customization unless they add to these macros.
Private Sub ckCodingCF()
'conditional formating for ckCoding
  Call ckCoding.Cells.FormatConditions.Delete
    With ckCoding.Range("ckCodingQTable").FormatConditions
         .Add Type:=xlExpression, Formula1:="=MOD($S2,2)"
'         .Item(.count).SetLastPriority
         With .Item(1).Interior
            .PatternColorIndex = xlAutomatic
            .Color = 16248029
            .TintAndShade = 0
         End With
         .Item(1).StopIfTrue = False
    End With
    
    With ckCoding.Range("ckCodingQTable[[nRev]:[Nunrlsbl]]").FormatConditions
    .Add Type:=xlExpression, Formula1:="=AND($T2<""M"",$AL2<$AM2+3)"
    .Item(.count).SetFirstPriority
    With .Item(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 14737401
        .TintAndShade = 0
    End With
    .Item(1).StopIfTrue = False
    End With
    
    With ckCoding.Range("ckCodingQTable[[pgm_annc_id]:[PO]]").FormatConditions
    .Add Type:=xlExpression, Formula1:="=AND($T2=""N"", C2<>C1)"
    .Item(.count).SetFirstPriority
    With .Item(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 13630971
        .TintAndShade = 0
    End With
    .Item(1).StopIfTrue = False
    End With
    
    With ckCoding.Range("ckCodingQTable[prop_titl_txt]").FormatConditions
      .Add Type:=xlExpression, Formula1:="=$AU2"
      .Item(.count).SetFirstPriority
    With .Item(1).Interior
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent2
        .TintAndShade = 0.799981688894314
    End With
    .Item(1).StopIfTrue = False
    End With
    
    With ckCoding.Range("ckCodingQTable[[bas_rsch_pct]:[other_pct]]").FormatConditions
      .Add Type:=xlExpression, Formula1:="=($J2+$I2)<>1"
    .Item(.count).SetFirstPriority
    With .Item(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 14540024
        .TintAndShade = 0
    End With
    .Item(1).StopIfTrue = False
    End With
    
   With ckCoding.Range("ckCodingQTable[[st_code]:[wmd]]").FormatConditions
    .Add Type:=xlExpression, Formula1:="=O2"
    .Item(.count).SetFirstPriority
    With .Item(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 11518971
        .TintAndShade = 0
    End With
    .Item(1).StopIfTrue = False
   End With
End Sub

Private Sub ckSplitsCF()
'conditional formating for ckSplits
  Call ckSplits.Cells.FormatConditions.Delete
    With ckSplits.Range("ckSplitsQTable").FormatConditions
         .Add Type:=xlExpression, Formula1:="=MOD($S2,2)"
'         .Item(.count).SetLastPriority
         With .Item(1).Interior
            .PatternColorIndex = xlAutomatic
            .Color = 13497835
            .TintAndShade = 0
         End With
         .Item(1).StopIfTrue = False
    End With
    
    With ckSplits.Range("ckSplitsQTable").FormatConditions
    .Add Type:=xlExpression, Formula1:="=$U1<>$U2"
    .Item(.count).SetFirstPriority
    With .Item(1).Borders(xlTop)
        .LineStyle = xlContinuous
        .TintAndShade = 0
        .Weight = xlThin
    End With
    .Item(1).StopIfTrue = False
    End With
    
    With ckSplits.Range("ckSplitsQTable[bObj],ckSplitsQTable[bOrg],ckSplitsQTable[bPEC],ckSplitsQTable[bPO],ckSplitsQTable[bPRCs]").FormatConditions
    .Add Type:=xlExpression, Formula1:="=E2<>F2"
    .Item(.count).SetFirstPriority
    With .Item(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 11518971
        .TintAndShade = 0.55
    End With
    .Item(1).StopIfTrue = False
    End With
End Sub

Private Sub ckAwdCF()
'conditional formating for ckAwd
  Call ckAwd.Cells.FormatConditions.Delete
  ckAwd.Range("ckAwdQTable").RowHeight = 15 ' don't let abstracts mess up row height
    With ckAwd.Range("ckAwdQTable").FormatConditions
         .Add Type:=xlExpression, Formula1:="=MOD($M2,2)"
'         .Item(.count).SetLastPriority
         With .Item(1).Interior
            .PatternColorIndex = xlAutomatic
            .Color = 13497835
            .TintAndShade = 0
         End With
         .Item(1).StopIfTrue = False
    End With
    With ckAwd.Range("ckAwdQTable[[pgm_annc_id]:[cntx_stmt_id]]").FormatConditions
        .Add Type:=xlExpression, Formula1:="=AND($N2=""N"",C2<>C1)"
        .Item(.count).SetFirstPriority
        With .Item(1).Interior
           .PatternColorIndex = xlAutomatic
           .Color = 11518971
           .TintAndShade = 0.55
        End With
        .Item(1).StopIfTrue = False
    End With
    With ckAwd.Range("ckAwdQTable[[rqst_eff_date]:[Country]]").FormatConditions
        .Add Type:=xlExpression, Formula1:="=AND($N2=""N"",Trim(X2)<>Trim(X1))"
        .Item(.count).SetFirstPriority
        With .Item(1).Interior
           .PatternColorIndex = xlAutomatic
           .Color = 11518971
           .TintAndShade = 0.55
        End With
        .Item(1).StopIfTrue = False
    End With
    With ckAwd.Range("ckAwdQTable[[prop_titl_txt]]").FormatConditions
        .Add Type:=xlExpression, Formula1:="=$BC2"
        .Item(.count).SetFirstPriority
        With .Item(1).Interior
            .PatternColorIndex = xlAutomatic
            .Color = 14540024
            .TintAndShade = 0
        End With
        .Item(1).StopIfTrue = False
    End With
End Sub
