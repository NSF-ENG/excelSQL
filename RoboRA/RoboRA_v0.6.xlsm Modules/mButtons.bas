Attribute VB_Name = "mButtons"
Option Explicit

Sub pidsFromTables()
Dim SQL As String
SQL = "CREATE TABLE #myPids (prop_id char(7) primary key, RAtemplate varchar(63))" & vbNewLine _
& IDsFromColumnRange("INSERT INTO  #myPids SELECT prop_id, '" & _
        RoboRA.Range("AwdTemplate") & "' as RAtemplate FROM csd.prop WHERE prop_id IN ", _
        "AwdPropTable[[prop_id]]") _
& IDsFromColumnRange("INSERT INTO  #myPids SELECT prop_id, '" & _
        RoboRA.Range("DeclTemplate") & "' as RAtemplate FROM csd.prop WHERE prop_id IN ", _
        "DeclPropTable[[prop_id]]") _
& IDsFromColumnRange("INSERT INTO  #myPids SELECT prop_id, '" & _
        RoboRA.Range("StdDeclTemplate") & "' as RAtemplate FROM csd.prop WHERE prop_id IN ", _
        "StdDeclPropTable[[prop_id]]")
        Debug.Print SQL
End Sub


Sub OptionButton_AreYouSure()
  If MsgBox("Are you sure that you want to overwrite RAs that may exist in eJacket?", _
            vbOKCancel) <> vbOK Then RoboRA.Range("overwrite_option").Value = 2
End Sub

Sub RefreshFromPanel()

Dim panl_id As String
panl_id = Replace(Replace(Advanced.Range("panl_id"), " ", ""), ",", "','")

Dim myProps As String
myProps = propHeader() _
& "JOIN csd.panl_prop pp ON p.prop_id = pp.prop_id" & vbNewLine _
& "WHERE p.prop_stts_code IN ('00','01','02','08','09') AND pp.panl_id in ('" & panl_id & "')" & vbNewLine _
& propTrailer()

Call doQuery(myProps)
End Sub

Sub RefreshFromBlock()

Dim org_code As String
org_code = Advanced.Range("org_code")
Dim pgm_ele_code As String
pgm_ele_code = Advanced.Range("pgm_ele_code")
Dim pm_ibm_logn_id As String
pm_ibm_logn_id = Advanced.Range("pm_ibm_logn_id")
Dim rcvd_before As String
rcvd_before = Format(Advanced.Range("rcvd_before"), "yyyy-mm-dd hh:mm:ss")
Dim solicitation As String
solicitation = Advanced.Range("solicitation")

Dim myProps As String
myProps = propHeader() _
& "WHERE p.prop_stts_code IN ('00','01','02','08','09')" & vbNewLine _
& "AND (p.pgm_annc_id like '" & solicitation & "') AND (p.org_code like '" & org_code & "') " & vbNewLine _
& "AND (p.pgm_ele_code like '" & pgm_ele_code & "') AND (p.pm_ibm_logn_id like '" & pm_ibm_logn_id & "') " & vbNewLine _
& "AND (p.nsf_rcvd_date < {ts '" & rcvd_before & "'}) " & vbNewLine _
& propTrailer()

Call doQuery(myProps)
End Sub

Sub List_Templates() ' list RA templates available (used by data validation)
Dim templateName As String
Dim nTemplates As Integer
nTemplates = 0

With Sheets("Advanced").ListObjects("AvailableTemplates")
  If .ListRows.count > 0 Then .DataBodyRange.Delete
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
End Sub

Sub Picker_dirRAtemplate()
Range("dirRAtemplate").Value = FolderPicker("Choose input folder containing RA templates", Range("dirRAtemplate").Value)
Call List_Templates
End Sub

Sub Picker_dirRAoutput()
  Range("dirRAoutput").Value = FolderPicker("Choose output folder for populated RAs (.docm)", Range("dirRAoutput").Value)
End Sub

Sub Install_Raddin()
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