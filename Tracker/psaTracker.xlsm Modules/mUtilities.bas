Attribute VB_Name = "mUtilities"
Option Explicit
Sub CleanUpSheet(ws As Worksheet, Optional emptyRow As Long = 7)
' delete blank rows on sheet below lowest listObject range
' pass emptyrow as the index of the first row that could be empty.
Dim i, r, lastRow As Long
With ws
    On Error Resume Next
    For i = 1 To ws.ListObjects.count
      r = ws.ListObjects(i).Range.End(xlDown).Row + 1
      If emptyRow < r Then emptyRow = r
    Next i
    
    For i = 1 To ws.PivotTables.count
      r = ws.PivotTables(i).RowRange.End(xlDown).Row + 1
      If emptyRow < r Then emptyRow = r
    Next i
    On Error GoTo 0
    lastRow = ws.UsedRange.Rows.count
    If emptyRow < lastRow Then ws.Rows(emptyRow & ":" & lastRow).Delete
End With
End Sub

Sub RefreshPivotTables(ws As Worksheet, QT As QueryTable)
 Dim PT As PivotTable
 For Each PT In ws.PivotTables
   PT.PivotTableWizard SourceType:=xlDatabase, SourceData:=QT.ListObject.Name
   If Not (PT Is Nothing) Then PT.RefreshTable
  Next
End Sub

Sub ClearTable(LO As ListObject)
  With LO
    If Not .DataBodyRange Is Nothing Then .DataBodyRange.Delete
  End With
End Sub
