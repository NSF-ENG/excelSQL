Attribute VB_Name = "mQueries"
Option Explicit

Sub BasicQueries(myPids As String)
With HiddenSettings
    'need this first to get revtable in order
    Call doQuery(PRCs.ListObjects("PECGlossaryTable").QueryTable, myPids _
        & .Range("RA_PECglossary") & .Range("revtable") & "DROP TABLE #myPid")

    Call doQuery(PRCs.ListObjects("PRCGlossaryTable").QueryTable, myPids _
        & .Range("RA_leads") & .Range("RA_propPRCs") & .Range("RA_PRCglossary") _
        & "DROP TABLE #myPid, #myLead, #myRA DROP TABLE #myPRCs, #myPRCdata")
    
    ' this is the slowest; do it in the background
    Call doQuery(ProjText.ListObjects("ProjTextTable").QueryTable, myPids _
        & .Range("RA_leads") & .Range("RA_projText") _
        & "DROP TABLE #myPid, #myLead, #myRA DROP TABLE #myRevInfo, #mySumm")
        
    Call doQuery(ckCoding.ListObjects("ckCodingTable").QueryTable, myPids _
        & .Range("RA_leads") & .Range("RA_propPRCs") & .Range("RA_revs") _
        & .Range("RA_prop") & .Range("RA_panl") & .Range("RA_propCheck") _
        & "DROP TABLE #myPid, #myLead, #myRA, #myPRCs, #myPRCdata DROP TABLE #myRevs, #myRevPanl, #myRevMarks, #myRevSumm DROP TABLE #myPropBudg, #myProp DROP TABLE #myPanl, #myProjPanl, #myProjPanlSumm")
    
    Call doQuery(RAData.ListObjects("RADataTable").QueryTable, myPids _
        & .Range("RA_leads") & .Range("RA_propPRCs") & .Range("RA_revs") _
        & .Range("RA_prop") & .Range("RA_panl") & .Range("RA_allRAdata") _
        & "DROP TABLE #myPid, #myLead, #myRA, #myProp, #myPropBudg, #myRevs, #myRevPanl, #myRevMarks, #myRevSumm, #myPanl, #myProjPanl, #myProjPanlSumm DROP TABLE #myDmog")
End With
Call CleanUpSheet(ckCoding)
Call CleanUpSheet(ProjText)
Call CleanUpSheet(RAData)
End Sub

Sub AwdCodingQueries(myPids As String)
With HiddenSettings
    'these can be done for awards only
    Call doQuery(Budgets.ListObjects("BudgetsTable").QueryTable, myPids _
        & .Range("RA_leads") & .Range("RA_propPRCs") & .Range("RA_budgBlocks") _
        & "DROP TABLE #myPid, #myLead, #myRA, #myPRCs, #myPRCdata DROP TABLE #myBudg")
        
    Call doQuery(ckAwd.ListObjects("ckAwdTable").QueryTable, myPids _
        & .Range("RA_leads") & .Range("RA_propPRCs") & .Range("RA_prop") _
        & .Range("RA_awdCheck") _
        & "DROP TABLE #myPid, #myLead, #myRA, #myPRCs, #myPRCdata DROP TABLE #myProp, #myPropBudg DROP TABLE #myCtry, #myCovrInfo, #myBudgPRC ")
    
    Call doQuery(ckSplits.ListObjects("ckSplitTable").QueryTable, myPids _
        & .Range("RA_leads") & .Range("RA_propPRCs") & .Range("RA_prop") & .Range("RA_splits") _
        & "DROP TABLE #myPid, #myLead, #myRA, #myPRCs, #myPRCdata DROP TABLE #myProp, #myPropBudg DROP TABLE #myBSprc")
End With
Call RefreshPivotTables(ckSplits)
Call CleanUpSheet(Budgets)
Call CleanUpSheet(ckSplits)
Call CleanUpSheet(ckAwd)
End Sub
