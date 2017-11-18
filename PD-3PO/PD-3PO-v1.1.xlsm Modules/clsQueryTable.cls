VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsQueryTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' This class is for refreshable sheets that have a query table.
' It requires each such sheet to have
'  1. dim QT as clsQueryTable
'  2. Sub InitializeQueryTable, called from ThisWorkbook_Open, instantiantes QT for the sheet
'  3. sheet-scoped named ranges run_datetime, run_rows where it can record results
'  4. RefreshPivotTables(ws) to refresh any pivot tables on sheet
'
' It supplies
'  1. initialize - helper for sheet's InitializeQueryTable
'  2. doQuery - helper for sheet's RefreshRS
'  3. handlers for events BeforeRefresh and AfterRefresh
'  4. clear - clear query table data body


Dim m_Sheet As Worksheet
Dim WithEvents m_QueryTable As QueryTable
Attribute m_QueryTable.VB_VarHelpID = -1

Public Property Get Sheet() As Worksheet
  Set Sheet = m_Sheet
End Property

Sub initialize(ws As Worksheet, QT As QueryTable)
'called once on ThisWorkbook.Open
    Set m_Sheet = ws
    Set m_QueryTable = QT
End Sub

Private Sub m_QueryTable_AfterRefresh(ByVal Success As Boolean)
' save datetime & number of rows in run_datetime and run_rows
 With m_Sheet.Range("run_datetime")
  If Success Then
    .Value = Format(Now, "'mm/dd/yy h:mm:ss am/pm")
  Else
    .Value = "unsuccessful"
  End If
 End With
 With m_Sheet.Range("run_nrows")
  If m_QueryTable.ListObject.DataBodyRange Is Nothing Then
    .Value = 0
  Else
    .Value = m_QueryTable.ListObject.DataBodyRange.Rows.Count
  End If
 End With
 Call RefreshPivotTables(m_Sheet) ' refresh any pivot tables on this sheet
 Call CleanupSheet(m_Sheet)
End Sub

Private Sub m_QueryTable_BeforeRefresh(Cancel As Boolean)
  m_Sheet.Range("run_datetime") = "query..."
End Sub

Sub clear()
' clear data from query table
  Call ClearTable(m_QueryTable.ListObject)
  Call CleanupSheet(m_Sheet)
  m_Sheet.Range("run_nrows").Value = 0
  m_Sheet.Range("run_datetime").Value = Format(Now(), "'mm/dd/yy h:mm:ss am/pm")
End Sub

' doQuery is to be called at the end of every query worksheet's RefreshRS.
' it assumes that the query table is defined
' It tries to handle errors, rerunning the query on a timeout and allowing users to cancel if they wish.

Sub doQuery(queryString As String)
' Worksheet's RefreshRS(limit) can call this function with something like:
'             QT.doQuery(mInputTab.InputSQL() & mainQuery & dropTables)

Dim i As Long
Dim errCount As Long
errCount = 0
    On Error GoTo cantRefresh:
    m_QueryTable.CommandText = queryString
    On Error GoTo catchTimeout
    m_QueryTable.Refresh '(False)
    DoEvents
cleanExit:
    On Error GoTo 0
    'Application.StatusBar = "done " & mySheet.Name
    'DoEvents
    Exit Sub
   
cantRefresh:
    If MsgBox("Error " & Err.Number & ": Cannot find query table on sheet " _
        & m_Sheet.name & "; skipping. " & Err.description, vbOKCancel) <> vbOK Then End
    Resume cleanExit
    
catchTimeout:
    If Err.Number = 1004 Then 'probably odbc error; maybe timeout
       errCount = errCount + 1
       If errCount < 2 Then Resume ' retry query once in case we timed out
    End If
    If MsgBox("Error " & Err.Number & ": cannot complete query on sheet " _
        & m_Sheet.name & "; skipping. " & Err.description, vbOKCancel) <> vbOK Then End
    Resume cleanExit
End Sub



