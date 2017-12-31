Attribute VB_Name = "mButtons"
Option Explicit

Sub ClearData()
 frmClearProp.Show
End Sub

Sub setPrefs()

End Sub

Sub OptionButton_AreYouSure()
  If MsgBox("Are you sure that you want to overwrite RAs that may exist in eJacket?", _
            vbOKCancel) <> vbOK Then Range("overwrite_option").Value = 2
End Sub

Sub copyLocationAndClose()
' if opened on a mac
CopyText (ThisWorkbook.FullName)
#If Mac Then
ThisWorkbook.Close savechanges:=False
#Else
If MsgBox("This is for mac users; I've copied the location; do you really want me to close without saving?", vbOKCancel) <> vbOK Then End
#End If
End Sub

Sub PullDataFromTables()
Dim sql2 As String
Dim awdSQL As String
Dim allSQL As String
' Query pulling from tables
sql2 = "' as RAtemplate FROM csd.prop p WHERE prop_stts_code like '" _
        & Advanced.Range("prop_stts_code") & "' AND prop_id IN "

With HiddenSettings
 awdSQL = IDsFromColumnRange("INSERT INTO #myPidRAt " & .Range("RA_pidRAtSelect") _
        & "'" & RoboRA.Range("AwdTemplate") & sql2, "AwdPropTable[[prop_id]]")
 allSQL = awdSQL & IDsFromColumnRange("INSERT INTO #myPidRAt " & .Range("RA_pidRAtSelect") _
        & "'" & RoboRA.Range("DeclTemplate") & sql2, "DeclPropTable[[prop_id]]") _
    & IDsFromColumnRange("INSERT INTO #myPidRAt " & .Range("RA_pidRAtSelect") _
        & "'" & RoboRA.Range("StdDeclTemplate") & sql2, "StdDeclPropTable[[prop_id]]") _
    & IDsFromColumnRange("INSERT INTO #myPidRAt " & .Range("RA_pidRAtSelect") _
        & "'" & RoboRA.Range("StdNDPDeclTemplate") & sql2, "StdNDPDeclPropTable[[prop_id]]")
 Call BasicQueries(.Range("RA_pidRAt") & allSQL)
 Call AwdCodingQueries(.Range("RA_pidRAt") & awdSQL)
End With
End Sub


Sub RefreshFromBlock()
' Advanced query with parameters from PD-3PO like block
  mySQLFrom = "INTO #myPidRAt FROM csd.prop prop" & vbNewLine
  mySQLWhere = ""
  With Advanced
    If hasValue("from_date") Then mySQLWhere = mySQLWhere & "AND prop.nsf_rcvd_date >= {ts '" & VBA.Format$(.Range("from_date"), "yyyy-mm-dd hh:mm:ss") & "'} " & vbNewLine
    If hasValue("to_date") Then mySQLWhere = mySQLWhere & "AND prop.nsf_rcvd_date <= {ts '" & VBA.Format$(.Range("to_date"), "yyyy-mm-dd hh:mm:ss") & "'} " & vbNewLine
  End With
  Call whereField("pgm_annc_id")
  Call whereField("org_code")
  Call whereField("pgm_ele_code")
  Call whereField("obj_clas_code")
  Call whereField("prop_titl_txt")
  Call whereField("pm_ibm_logn_id")
  Call whereField("dir_div_abbr", "org", "_code")
  Call whereField("panl_id", "panl_prop", "prop_id")
  Call whereField("_code", "prop_atr", "prop_id", notPreamble:=" AND prop_atr.prop_atr_type_code = 'PRC'")
  If Len(mySQLWhere) < 3 Then
    MsgBox ("Please restrict the set of proposals by panel, solicitation, PD, or something. Exiting.")
    End
  End If
  Call whereField("prop_stts_code")
  Call whereField("_abbr", "natr_rqst", "_code")
  mySQLWhere = "WHERE (1=1) " & mySQLWhere
  
  Dim query As String
  With HiddenSettings
   query = "SET NOCOUNT ON" & vbNewLine & .Range("RA_pidSelect") & "convert(varchar(63),'') as RAtemplate " & vbNewLine & mySQLFrom & mySQLWhere
   Call BasicQueries(query)
   Call AwdCodingQueries(query)
  End With
End Sub

Sub pasteFromMyWork()
' Meant for copying proposal ids from the eJ mywork screen,
' this gets anything that is 5-7 digit number in the first 10 characters of each line in the clipboard.
' and pastes as a column down from the active cell
Dim i As Long, p As Long
Dim txt() As String
Dim prop_id As String
Dim cb As DataObject
#If Mac Then
    Set cb = New DataObject
#Else
    Set cb = CreateObject("new:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")
#End If
cb.GetFromClipboard
On Error Resume Next
prop_id = cb.GetText()
On Error GoTo 0
If prop_id <> "" Then
    txt = Split(prop_id, vbLf)
    prop_id = ""
    For i = LBound(txt) To UBound(txt)
      p = Val(Left(txt(i), 10))
    '  Debug.Print i & Left(txt(i), 10) & p
      If (p >= 100000# And p < 100000000#) Then prop_id = prop_id & Format(p, "0000000") & vbLf ' a hack to recognize prop_ids
    Next
End If
If Len(prop_id) < 8 Then
 MsgBox ("No prop_ids in clipboard. Please copy your eJacket MyWork page, select a cell in any prop_id table, then click Paste From MyWork.")
 Exit Sub
End If
cb.Clear
If MsgBox("OK to paste these proposal ids starting in cell " & Selection.Address & "?" & vbNewLine & prop_id, vbOKCancel) = vbOK Then
    txt = Split(prop_id, vbLf)
    Range(ActiveCell, ActiveCell.Offset(UBound(txt) - LBound(txt))).Value = Application.Transpose(txt)
    HiddenSettings.Range("select_prop_stts").Value = 3 ' set to DD_concur
End If
End Sub

Sub testPermissions()
' it turns out that the only non-public table we use is csd.pgm_ref
' If we can't access it, then limit the PRCGlossary query
Dim rtn As Long
Call handlePwd
rtn = tryConnection("select * from csd.pgm_ref where pgm_ref_code = '7929'")
If rtn = 0 Then ' can access pgm_ref
  Prefs.Range("test_table_permissions").Value = "Permissions tested OK, " & VBA.Format(Now(), "Medium Date")
ElseIf MsgBox("Accessing table rptdb.csd.pgm_ref got this response:" & vbNewLine & lastConnectionErrorDescription & vbNewLine _
& "I'll disable program reference code names in the PRC Glossary so you can still use RoboRA", vbOKCancel) = vbOK Then
  Prefs.Range("test_table_permissions").Value = "rptdb.csd.pgm_ref table removed since access unavailable " & VBA.Format(Now(), "Medium Date")
  HiddenSettings.Range("RA_PRCGlossary").Value = HiddenSettings.Range("RA_no_PRC").Value
End If
End Sub
