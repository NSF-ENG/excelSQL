Attribute VB_Name = "Module_Buttons"
Sub sortByPanlRec()
'
' sortByPanlRec Macro
'
With ActiveWorkbook.Sheets("PropsOnPanls").ListObjects("PropPanelTable").Sort
    With .SortFields
        .clear
        .Add Key:=Range("PropPanelTable[rcom_seq]"), SortOn:=xlSortOnValues, order:=xlAscending, DataOption:=xlSortNormal
        .Add Key:=Range("PropPanelTable[rank]"), SortOn:=xlSortOnValues, order:=xlAscending, DataOption:=xlSortNormal
        .Add Key:=Range("PropPanelTable[avg_score]"), SortOn:=xlSortOnValues, order:=xlDescending, DataOption:=xlSortNormal
        .Add Key:=Range("PropPanelTable[lead]"), SortOn:=xlSortOnValues, order:=xlAscending, DataOption:=xlSortTextAsNumbers
        .Add Key:=Range("PropPanelTable[panl_id]"), SortOn:=xlSortOnValues, order:=xlAscending, DataOption:=xlSortNormal
        .Add Key:=Range("PropPanelTable[ILN]"), SortOn:=xlSortOnValues, order:=xlAscending, DataOption:=xlSortNormal
        .Add Key:=Range("PropPanelTable[prop_id]"), SortOn:=xlSortOnValues, order:=xlAscending, DataOption:=xlSortTextAsNumbers
    End With
    .Header = xlYes
    .MatchCase = False
    .Orientation = xlTopToBottom
    .SortMethod = xlPinYin
    .Apply
End With
Range("PropPanelTable[[#Headers],[rcom_abbr]]").Activate
End Sub
Sub SortbyPanlDate()
'
' SortbyPanlDate Macro
'

With ActiveWorkbook.Sheets("PropsOnPanls").ListObjects("PropPanelTable").Sort
        With .SortFields
            .clear
            .Add Key:=Range("PropPanelTable[panl_bgn_date]"), SortOn:=xlSortOnValues, order:=xlAscending, DataOption:=xlSortNormal
            .Add Key:=Range("PropPanelTable[panl_id]"), SortOn:=xlSortOnValues, order:=xlAscending, DataOption:=xlSortNormal
            .Add Key:=Range("PropPanelTable[disc_ordr]"), SortOn:=xlSortOnValues, order:=xlAscending, DataOption:=xlSortNormal
            .Add Key:=Range("PropPanelTable[lead]"), SortOn:=xlSortOnValues, order:=xlAscending, DataOption:=xlSortTextAsNumbers
            .Add Key:=Range("PropPanelTable[ILN]"), SortOn:=xlSortOnValues, order:=xlAscending, DataOption:=xlSortNormal
            .Add Key:=Range("PropPanelTable[prop_id]"), SortOn:=xlSortOnValues, order:=xlAscending, DataOption:=xlSortTextAsNumbers
        End With
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Range("PropPanelTable[[#Headers],[disc_ordr]]").Activate
End Sub

Sub SendCOI()
  ' construct email to COI checker
  Dim mailBody As String
  Dim i, nprop As Long
  Dim nrevr, estHours As Double
  With BulkCOI.Range("COIPropTable[prop_id]") ' list prop ids in mail message
    nprop = .Rows.Count
    If nprop < 2 Then
      If nprop = 0 Or Len(.Value) <> 7 Then
        MsgBox "No proposals listed; aborting."
        Exit Sub
      End If
      mailBody = .Value
     Else ' we have at least two, and can use transpose
      mailBody = Join(Application.Transpose(.Value), vbNewLine)
    End If
  End With
  mailBody = mailBody & vbNewLine & vbNewLine
  With BulkCOI.Range("COIRevrTable[[whole name]:[email]]") 'list reviewers in mail message
   nrevr = .Rows.Count
   If nrevr = 0 Then
        MsgBox "No reviewers listed; aborting."
        Exit Sub
   End If
   estHours = 0.1 + nprop / 90
   For i = 1 To nrevr
    mailBody = mailBody & Replace(.Cells(i, 1), ",", " ") & ", " & .Cells(i, 2) & vbNewLine
   Next i
  End With
  
  If MsgBox("You are requesting that MPS Proposal Check compare " & nprop & " projects againat " & nrevr & " reviewer names, and email a single results spreadsheet." & vbNewLine & _
   "Expect this to take about " & Format(estHours / 24, "Short Time") & "(depending on the server load), so look for the email around " & Format(Now() + estHours / 24, "medium time") & ".", vbOKCancel) <> vbOK Then Exit Sub
   If nprop > 100 Then
       If (Hour(Now()) > 5 And Hour(Now()) < 18) Then If MsgBox("Are you sure you want the results of checking all proposals on " & nprop & " projects now? (Jobs of this size should be done overnight.)", vbOKCancel) <> vbOK Then Exit Sub
       If nprop > 500 Then If MsgBox("Last chance: You are about to commit significant resources to downloading all sections of " & nprop & " projects.", vbOKCancel) <> vbOK Then Exit Sub
   End If
  Call SendMail("proposal@nsf.gov", "COI [MERGECONFLICTONLY]", mailBody) ' send mail
End Sub

Sub GetAssignedRevr()
  ' construct email to COI checker
Dim queryString As String
If BulkCOI.Range("COIPropTable").Rows.Count < 1 Then
    MsgBox "No proposals listed; aborting"
    Exit Sub
End If

If BulkCOI.Range("COIRevrTable").Rows.Count > 1 Then If MsgBox("Confirm that you want to replace the current list of reviewers with the assigned reviewers.", vbOKCancel) <> vbOK Then Exit Sub
    

queryString = "SELECT DISTINCT revr.revr_last_name, revr.revr_frst_name, a.revr_addr_txt AS 'email'" & vbNewLine _
& "FROM csd.rev_prop rp" & vbNewLine _
& "JOIN csd.revr revr ON rp.revr_id = revr.revr_id" & vbNewLine _
& "LEFT OUTER JOIN csd.revr_opt_addr_line a ON rp.revr_id = a.revr_id AND a.addr_lne_type_code='E'" & vbNewLine _
& IDsFromColumnRange(" WHERE rp.prop_id IN", BulkCOI.Range("COIPropTable[prop_id]")) _
& " ORDER BY revr.revr_last_name, revr.revr_frst_name" & vbNewLine

With BulkCOI.ListObjects("COIRevrTable").QueryTable
  .CommandText = queryString
  .Refresh False
End With
End Sub

