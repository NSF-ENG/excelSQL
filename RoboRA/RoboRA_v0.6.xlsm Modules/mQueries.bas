Attribute VB_Name = "mQueries"
Option Explicit
Sub makeQueries(myPids As String, Optional myAwds As String = "")
If Len(myAwds) < 2 Then myAwds = myPids

With HiddenSettings
    myPids = .Range("RA_pidSelect") & myPids
    myAwds = .Range("RA_pidSelect") & myAwds
    'need this first to get revtable in order
    Call doQuery(PRCs.ListObjects("PRCGlossaryTable").QueryTable, myPids & .Range("revtable") _
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
    
    'these can be done for awards only
    Call doQuery(.ListObjects("BudgetsTable").QueryTable, myAwds _
        & .Range("RA_leads") & .Range("RA_budgBlocks") _
        & "DROP TABLE #myPid, #myLead, #myRA DROP TABLE ")
        
    Call doQuery(ckAwd.ListObjects("ckAwdTable").QueryTable, myAwds _
        & .Range("RA_leads") & .Range("RA_propPRCs") & .Range("RA_prop") _
        & .Range("RA_revs") & .Range("RA_awdCheck") _
        & "DROP TABLE #myPid, #myLead, #myRA DROP TABLE ")
    
    Call doQuery(ckSplits.ListObjects("ckSplitTable").QueryTable, myAwds _
        & .Range("RA_leads") & .Range("RA_propPRCs") & .Range("RA_prop") & .Range("RA_splits") _
        & "DROP TABLE #myPid, #myLead, #myRA DROP TABLE #myProp, #myPropBudg DROP TABLE #myBudgPRC")
End With

Call CleanUpSheet(ckCoding)
Call CleanUpSheet(ckCoding)
Call CleanUpSheet(ckAwd)
Call CleanUpSheet(ProjText)
Call CleanUpSheet(Budgets)
Call CleanUpSheet(RAData)
Call CleanUpSheet(ckSplits)
End Sub
