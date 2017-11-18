Attribute VB_Name = "Module_UtilityFcns"
'Jack Snoeyink Oct 2016
' Utility functions for strings from ranges, and clearing sheets

Option Explicit

Sub SendMail(strTo As String, strSubject As String, strBody As String, Optional strCC As String, _
                    Optional strBCC As String, Optional strAttachment As String)
   'http://excelribbon.tips.net/T011785_Automatic_Text_in_an_E-mail.html
   
    Dim OutApp As Object
    Dim OutMail As Object

    Set OutApp = CreateObject("Outlook.Application")
    Set OutMail = OutApp.CreateItem(0)

    With OutMail
        .To = strTo
        .Subject = strSubject
    On Error Resume Next ' skip the omitted
        .CC = strCC
        .BCC = strBCC
        .Body = strBody
        .Attachments.Add (strAttachment)
     On Error GoTo SendMailErr
        .Send 'or use .Display
    End With

SendMailExit:
    On Error GoTo 0
    Set OutMail = Nothing
    Set OutApp = Nothing
    Exit Sub
SendMailErr:
    MsgBox ("Error " & Err.Number & " attempting to send mail: " & Err.description)
    Resume SendMailExit
End Sub


Function IDsFromColumnRange(prefix As String, colRange As Range) As String
' returns the set of column entries, stripped of spaces, suitable for an IN clause or "" if none.
Dim ids As String
If colRange.Rows.Count < 2 Then
    ids = colRange.Value
Else ' we have at least two, and can use transpose
    ids = Join(Application.Transpose(colRange.Value), "','") ' quote column entries and make a comma-separated row
End If
ids = "'" & Replace(Replace(ids, " ", ""), Chr(160), "") & "'" ' strip spaces (visible and invisible) and ,'' from string.
ids = Replace(ids, ",''", "")  ' strip blank column entries
'MsgBox "please check your ids : " + ids
If Len(ids) < 3 Then
    IDsFromColumnRange = ""
Else
    IDsFromColumnRange = prefix & " (" & ids & ")" & vbLf
End If
End Function

Function clipboardRangeAsCSV()
' return clipboard as comma-separated values (CSV) -- tabs and newlines replaced by commas.
 Dim DataObj As New MSForms.DataObject
 Dim s As String
 Dim i As Long
    DataObj.GetFromClipboard
    s = Replace$(Replace$(DataObj.GetText, vbTab, ","), vbNewLine, ",")
    i = Len(s)
    While i > 0 And Mid$(s, i, 1) = ","
      i = i - 1
    Wend
    clipboardRangeAsCSV = Left$(s, i)
End Function

Sub pasteRangeAsCSV()
' paste comma-separated range from clipboard into active cell, after confirmation.
Dim cboard As String
On Error GoTo pasteRangeAsCSVError
    cboard = clipboardRangeAsCSV()
    If MsgBox("Ok to paste " & cboard & " into " & ActiveCell.Address & "?", vbOKCancel) <> vbOK Then End
    ActiveCell.Value = cboard
pasteRangeAsCSVExit:
    On Error GoTo 0
    Exit Sub
pasteRangeAsCSVError:
    MsgBox "Need a range of cells copied to clipboard and an active cell to paste into. " _
      & Err.Number & Err.description
    Resume pasteRangeAsCSVExit
End Sub

Sub ClearSortAndFilter()
' clears sort and filter from table on active sheet ; from macro recorder -- add error handling
'On Error Resume Next
With ActiveSheet.ListObjects(1)
    .DataBodyRange.AutoFilter ' clear and restore
    .DataBodyRange.AutoFilter
    .Sort.SortFields.clear
End With
On Error GoTo 0
End Sub

Sub RefreshPivotTables(ws As Worksheet)
' Assumes ws is a RefreshableSheet with required named ranges
' JSS add error handling
    Dim pt As pivotTable
    For Each pt In ws.PivotTables
       On Error Resume Next
       pt.RefreshTable
       If Err.Number <> 0 Then MsgBox "Can't refresh pivot table " & pt.name & " on " & ws.name & ". Skipping."
       
       DoEvents
       ws.Range("run_datetime").Value = Format(pt.RefreshDate, "'mm/dd/yy h:mm:ss am/pm")
    Next
    On Error GoTo 0
End Sub

Sub ClearTable(lo As ListObject)
 With lo ' clear table DataBodyRange
    If Not .DataBodyRange Is Nothing Then .DataBodyRange.Delete
 End With
End Sub

Sub CleanupSheet(ws As Worksheet)
' delete blank rows below lowest listObject range
Dim i, r, emptyRow, lastRow As Long
emptyRow = 7
With ws
    On Error Resume Next
    For i = 1 To .ListObjects.Count
      r = .ListObjects(i).Range.End(xlDown).Row + 1
      If emptyRow < r Then emptyRow = r
    Next i
    For i = 1 To .PivotTables.Count
      r = .PivotTables(i).RowRange.End(xlDown).Row + 1
      If emptyRow < r Then emptyRow = r
    Next i
    On Error GoTo 0
    lastRow = .UsedRange.Rows.Count
    If emptyRow < lastRow Then .Rows(emptyRow & ":" & lastRow).Delete
End With
End Sub
