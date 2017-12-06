Attribute VB_Name = "mButtons"
Option Explicit
Public Function SummarizeQuestMarks(abstr As String) As String
Dim i As Long
Dim s As String
s = " "
i = InStrRev(abstr, "?")
While i > 0
i = Application.Max(1, i - 3)
s = Mid(abstr, i, 7) & "|" & s
i = InStrRev(abstr, "?", i - 1)
Wend
SummarizeQuestMarks = s
End Function

Sub ClearQueryParams()
 Call List_Templates
 If MsgBox("Ok to clear query parameters?  (Can't undo)", vbOKCancel) <> vbOK Then End
 ActiveSheet.Range("query_params").Cells.Value = HiddenSettings.Range("query_params").Cells.Value
End Sub

Sub PullDataFromTables()
Dim PendOnly As String
Dim awdSQL As String
Dim allSQL As String
'PendOnly = "prop_stts_code BETWEEN '03' AND '0Z' AND "
'PendOnly = "prop_stts_code like ('0%') AND "
'PendOnly = "prop_stts_code like ('0[01289]') AND "
'PendOnly = "prop_stts_code like ('0[34]') AND "
PendOnly = "prop_stts_code IN ('00','01','02','08','09') AND "

With HiddenSettings
 awdSQL = IDsFromColumnRange("INSERT INTO #myPid " & .Range("RA_pidSelect") _
        & "'" & RoboRA.Range("AwdTemplate") & "' as RAtemplate FROM csd.prop p WHERE " _
        & PendOnly & "prop_id IN ", "AwdPropTable[[prop_id]]")
 allSQL = awdSQL & IDsFromColumnRange("INSERT INTO #myPid " & .Range("RA_pidSelect") _
        & "'" & RoboRA.Range("DeclTemplate") & "' as RAtemplate FROM csd.prop p WHERE " _
        & PendOnly & "prop_id IN ", "DeclPropTable[[prop_id]]") _
    & IDsFromColumnRange("INSERT INTO #myPid " & .Range("RA_pidSelect") _
        & "'" & RoboRA.Range("StdDeclTemplate") & "' as RAtemplate FROM csd.prop p WHERE " _
        & PendOnly & "prop_id IN ", "StdDeclPropTable[[prop_id]]")
 Call BasicQueries(.Range("RA_pidCreate") & allSQL)
 Call AwdCodingQueries(.Range("RA_pidCreate") & awdSQL)
End With
End Sub

Sub OptionButton_AreYouSure()
  If MsgBox("Are you sure that you want to overwrite RAs that may exist in eJacket?", _
            vbOKCancel) <> vbOK Then RoboRA.Range("overwrite_option").Value = 2
End Sub

Sub RefreshFromPanel()
'temporary, until we convert to parse
Dim panl_id As String
Dim pidWhere As String
panl_id = Replace(Replace(Advanced.Range("panl_id"), " ", ""), ",", "','")
pidWhere = "SET NOCOUNT ON" & vbNewLine _
& "SELECT DISTINCT p.prop_id, p.lead_prop_id, p.pi_id, p.inst_id, p.pm_ibm_logn_id as PO, convert(varchar(63),'') AS RAtemplate" & vbNewLine _
& "INTO #myPid FROM csd.prop p" & vbNewLine _
& "JOIN csd.panl_prop pp ON p.prop_id = pp.prop_id" & vbNewLine _
& "WHERE pp.panl_id in ('" & panl_id & "')" & vbNewLine
Call BasicQueries(pidWhere)
Call AwdCodingQueries(pidWhere)
End Sub

Sub RefreshFromBlock()
  mySQLFrom = "FROM csd.prop prop" & vbNewLine
  mySQLWhere = ""
  With Advanced
    If hasValue("from_date") Then mySQLWhere = mySQLWhere & "AND prop.nsf_rcvd_date >= {ts '" & Format(.Range("from_date"), "yyyy-mm-dd hh:mm:ss") & "'} " & vbNewLine
    If hasValue("to_date") Then mySQLWhere = mySQLWhere & "AND prop.nsf_rcvd_date <= {ts '" & Format(.Range("to_date"), "yyyy-mm-dd hh:mm:ss") & "'} " & vbNewLine
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
    Call whereField("_abbr", "prop_stts", "_code")
    Call whereField("_abbr", "natr_rqst", "_code")
    If Len(mySQLWhere) < 3 Then
        MsgBox ("Please restrict the set of proposals by panel, solicitation, PD, or something. Exiting.")
        End
    Else
        mySQLWhere = "WHERE (1=1) " & mySQLWhere
    End If

Dim query As String
With HiddenSettings
 query = "SET NOCOUNT ON" & .Range("RA_pidSelect") & "convert(varchar(63),'') as RAtemplate " & vbNewLine & mySQLFrom & mySQLWhere
 Call BasicQueries(query)
 Call AwdCodingQueries(query)
End With
End Sub

Sub List_Templates() ' list RA templates available (used by data validation)
Dim templateName As String
Dim nTemplates As Integer
nTemplates = 0
Application.ScreenUpdating = False
On Error GoTo ErrHandler
With Advanced.ListObjects("AvailableTemplates")
  If Not .DataBodyRange Is Nothing Then .DataBodyRange.Delete
  templateName$ = Dir(Range("dirRAtemplate").Value & "\*RAt.docx")
    Do While templateName$ <> ""
      If Left(templateName$, 1) <> "~" Then
        .ListRows.Add AlwaysInsert:=True
        nTemplates = nTemplates + 1
        .DataBodyRange(nTemplates, 1) = templateName$
        templateName$ = Dir
      End If
    Loop
End With
If nTemplates = 0 Then MsgBox ("Did not find any RA templates in " & Range("dirRAtemplate").Value & vbNewLine & "Remember that RA template names must end with RAt.docm")
ExitHandler:
Application.ScreenUpdating = True
Exit Sub
ErrHandler:
MsgBox ("Error " & Err.Number & ":" & Err.Description & vbNewLine & "while trying to list templates.  Ensure template directory, " & Range("dirRAtemplate").Value & ", is accessible.")
Resume ExitHandler
End Sub

Sub Picker_dirRAtemplate()
Range("dirRAtemplate").Value = FolderPicker("Choose input folder containing RA templates", Range("dirRAtemplate").Value)
Call List_Templates
End Sub

Sub Picker_dirRAoutput()
  Range("dirRAoutput").Value = FolderPicker("Choose output folder for populated RAs (.docm)", Range("dirRAoutput").Value)
End Sub

Sub installMacros()
Dim pathName As String

On Error GoTo ErrHandler:
'JSS PC vs mac version
pathName$ = "%appdata%\Microsoft\Word\STARTUP"
If Dir(pathName$, vbDirectory) = "" Then MkDir (pathName$)
MsgBox ("copying RAaddin.dotm into " & pathName$)
'JSS copy file RAaddin.dotm and trust it.
ExitHandler:
  On Error GoTo 0
  Exit Sub
ErrHandler:
  MsgBox ("Error in Install_Raddin: " & Err.Number & ":" & Err.Description)
  Resume ExitHandler
End Sub
