VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "BulkCOI"
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
     .description = "Check for CoI with potential reviewers"
     .tip = "Enter potential reviewers on tab " & .name
     .tabtitle = "Bulk COI Checking"
     .note = "Sends chosen projects and potential reviewers to MPS Proposal Check, which emails back a COI results spreadsheet. Expect the server to spend a minute per proposal. (Leads bring in collabs automatically.)"
     .helpText = "Constructs an Outlook email message to send chosen projects and potential reviewers to Paul Morris' MPS Proposal Check, which returns COI results in a spreadsheet by email to sender. This is an aid, not a cure-all; there will be false pasitives in name matches and false negatives from corrupted PDF documents (flagged) requiring manual scrutiny.  Response time depends on server load, but expect a minute per project."
     .helpvideo = ""
     .phase = 3
     .order = 3803
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
  Call ClearCOIRevr
  If Not QT Is Nothing Then Call QT.clear
End Sub

Sub RefreshRS(Optional limit As String = "")
'refresh tab based on parameters in InputTab
Dim dropTables As String
Dim mainQuery As String

'lists proposal ids
mainQuery = "SELECT Div, prop_id, ILN, pi_last_name, inst_name FROM #myProps prop WHERE ILN < 'M' AND prop_stts_abbr NOT IN ('WTH','RTNR') ORDER BY lead, ILN, prop_id" & vbNewLine

dropTables = "DROP TABLE #myProps"

Call QT.doQuery(mInputTab.InputSQL & mainQuery & dropTables)
End Sub


Sub ClearCOIRevr()
' Clear prospective reviewers and include whole name formula.
    Call ClearTable(Me.ListObjects("COIRevrTable"))
    Range("COIRevrTable[whole name]").FormulaR1C1 = "=[revr_frst_name] & "" "" &[revr_last_name]"
End Sub

Private Sub Worksheet_SelectionChange(ByVal Target As Range)

End Sub
