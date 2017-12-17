Attribute VB_Name = "mButtons"
Option Explicit


Sub ClearQueryParams()
 Call List_Templates
 If MsgBox("Ok to clear query parameters?  (Can't undo)", vbOKCancel) <> vbOK Then End
 ActiveSheet.Range("query_params").Cells.Value = HiddenSettings.Range("query_params").Cells.Value
End Sub

Sub OptionButton_AreYouSure()
  If MsgBox("Are you sure that you want to overwrite RAs that may exist in eJacket?", _
            vbOKCancel) <> vbOK Then RoboRA.Range("overwrite_option").Value = 2
End Sub

Sub PullDataFromTables()
Dim sql2 As String
Dim awdSQL As String
Dim allSQL As String
' Query pulling from tables
sql2 = "' as RAtemplate FROM csd.prop prop WHERE prop_stts_code like '" _
        & Advanced.Range("prop_stts_code") & "' AND prop_id IN "


With HiddenSettings
 awdSQL = IDsFromColumnRange("INSERT INTO #myPid " & .Range("RA_pidSelect") _
        & "'" & RoboRA.Range("AwdTemplate") & sql2, "AwdPropTable[[prop_id]]")
 allSQL = awdSQL & IDsFromColumnRange("INSERT INTO #myPid " & .Range("RA_pidSelect") _
        & "'" & RoboRA.Range("DeclTemplate") & sql2, "DeclPropTable[[prop_id]]") _
    & IDsFromColumnRange("INSERT INTO #myPid " & .Range("RA_pidSelect") _
        & "'" & RoboRA.Range("StdDeclTemplate") & sql2, "StdDeclPropTable[[prop_id]]")
 Call BasicQueries(.Range("RA_pidCreate") & allSQL)
 Call AwdCodingQueries(.Range("RA_pidCreate") & awdSQL)
End With
End Sub


Sub RefreshFromBlock()
' Advanced query with parameters from PD-3PO like block
  mySQLFrom = "INTO #myPid FROM csd.prop prop" & vbNewLine
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


'Sub installMacros()
'Dim pathName As String
'
'On Error GoTo ErrHandler:
''JSS PC vs mac version
'pathName$ = "%appdata%\Microsoft\Word\STARTUP"
'If Dir(pathName$, vbDirectory) = "" Then MkDir (pathName$)
'MsgBox ("copying RAaddin.dotm into " & pathName$)
''JSS copy file RAaddin.dotm and trust it.
'ExitHandler:
'  On Error GoTo 0
'  Exit Sub
'ErrHandler:
'  MsgBox ("Error in Install_Raddin: " & Err.Number & ":" & Err.Description)
'  Resume ExitHandler
'End Sub
