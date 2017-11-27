Attribute VB_Name = "mUtilities"
Option Explicit
Sub CleanUpSheet(ws As Worksheet, Optional emptyRow As Long = 7)
' delete blank rows on sheet below lowest listObject range
' pass emptyrow as the index of the first row that could be empty.
Dim i, r, lastRow As Long
With ws
    On Error Resume Next
    For i = 1 To ws.ListObjects.Count
      r = ws.ListObjects(i).Range.End(xlDown).Row + 1
      If emptyRow < r Then emptyRow = r
    Next i
    
    For i = 1 To ws.PivotTables.Count
      r = ws.PivotTables(i).RowRange.End(xlDown).Row + 1
      If emptyRow < r Then emptyRow = r
    Next i
    On Error GoTo 0
    lastRow = ws.UsedRange.Rows.Count
    If emptyRow < lastRow Then ws.Rows(emptyRow & ":" & lastRow).Delete
End With
End Sub

Sub RefreshPivotTables(ws As Worksheet, qt As QueryTable)
 Dim PT As PivotTable
 For Each PT In ws.PivotTables
   PT.PivotTableWizard SourceType:=xlDatabase, SourceData:=qt.ListObject.Name
   If Not (PT Is Nothing) Then PT.RefreshTable
  Next
End Sub

Sub ClearTable(lo As ListObject)
  With lo
    If Not .DataBodyRange Is Nothing Then .DataBodyRange.Delete
  End With
End Sub
Sub ClearMatchingTable(t As String)
' use wildcards to match table names to clear.
Dim lo As ListObject
For Each lo In ActiveSheet.ListObjects
  If (lo.Name Like t) Then Call ClearTable(lo)
Next
End Sub
Function FindTable(t As String) As ListObject
' find first table on active sheet matching pattern t
For Each FindTable In ActiveSheet.ListObjects
  If (FindTable.Name Like t) Then Exit Function
Next
Set FindTable = Nothing
End Function

