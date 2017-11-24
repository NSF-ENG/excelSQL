Attribute VB_Name = "Module1"
Option Explicit
Sub GetPI_AwdOblg()

Call handlePwd

Dim setNC, makeProps, prop_ids As String
Dim rng As Range
Dim i As Long

Set rng = Range("Table1[Proposal_IDs]")
rng.NumberFormat = "@" ' convert prop_ids in table to text for Match
For i = 1 To rng.Rows.count
  If IsNumeric(rng.Cells(i).Value) Then rng.Cells(i).Value = Format(rng.Cells(i).Value, "0000000")
Next i
' get list of prop_ids
If rng.Rows.count < 2 Then
    prop_ids = "= '" & rng.Value & "'" & vbNewLine
Else ' we have at least two, and can use transpose
    prop_ids = "In ('" & Join(Application.Transpose(rng.Value), "','") & "') " & vbNewLine
End If
prop_ids = Replace(Replace(prop_ids, ",''", ""), ",' '", "")

setNC = "SET NOCOUNT ON " & vbNewLine
makeProps = HiddenSettings.Range("proppiProps") & "WHERE prop.prop_id " & prop_ids _
             & "OR prop.lead_prop_id " & prop_ids
Call doQuery(Worksheets("Awards").ListObjects(1).QueryTable, _
             setNC & makeProps & HiddenSettings.Range("proppiAwd"))
Call doQuery(Worksheets("Pending").ListObjects(1).QueryTable, _
             setNC & makeProps & HiddenSettings.Range("proppiPend"))

Dim pivotTable As pivotTable
For Each pivotTable In Worksheets("Dashboard").PivotTables
   pivotTable.RefreshTable
Next
End Sub
