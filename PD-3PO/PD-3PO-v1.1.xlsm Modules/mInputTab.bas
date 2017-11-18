Attribute VB_Name = "mInputTab"
' Code for proposal-based input tab               Jack Snoeyink, Oct 2016
' Build the sql query to retrieve proposals based on csd.prop fields set on theis tab
' Database field names are used to name ranges in the spreadsheet where the use can set parameters
' Function andWhere supports parameter negation, lists, wildcards, and equals
' whereField can support simple joins to the csd.prop table
Option Explicit

' make these class variables because our subroutines below will create
'Dim InputTab As Worksheet
Dim myPidsWhere As String
Dim myPidsFrom As String


Private Function hasValue(rangeName As String) As Boolean
 With InputTab.Range(rangeName)
   hasValue = (Trim$(.Value) <> "" And Left$(.Value, 3) <> "eg:")
  End With
End Function

Function andWhere(tablename As String, fieldname As String)
' Field values have 4 cases:
'    nothing, AND field = value, AND field IN ('val1','val2',...,'valN'), _or_ AND field LIKE value
'    First character ~ negates any of the above conditions
'    Warning: trims spaces in value
Dim field As String
Dim optNeg As String
Dim hasComma As Boolean
Dim hasSqlWildcard As Boolean
optNeg = ""

field = Trim(InputTab.Range(fieldname).Value) ' Warning: trims spaces on values
If Left(field, 3) = "eg:" Then field = ""

If Left(field, 1) = "~" Then ' have negation
    optNeg = " NOT"
    field = Right(field, Len(field) - 1)
End If

hasComma = (InStr(field, ",") > 0) 'have list

hasSqlWildcard = (InStr(field, "%") > 0) Or (InStr(field, "_") > 0) Or ((InStr(field, "[") > 0) And (InStr(field, "]") > 0))
If (hasComma And hasSqlWildcard) Then
 MsgBox "Can't have SQL wildcards in a comma separated list for " & fieldname & ": " & field
 End
End If

If Len(field) < 1 Then ' do nothing
  andWhere = ""
ElseIf hasComma Then ' IN/NOT IN list
  andWhere = " AND " & tablename & fieldname & optNeg & " IN ('" & Replace(Replace(Join(Split(Replace(Replace(field, """", ""), "'", ""), ","), "','") & "')", " '", "'"), "' ", "'")
ElseIf hasSqlWildcard Then ' LIKE / NOT LIKE
  andWhere = " AND " & tablename & fieldname & optNeg & " LIKE '" & Replace(Replace(field, """", ""), "'", "") & "'"
ElseIf optNeg = "" Then ' equals
  andWhere = " AND " & tablename & fieldname & " = '" & field & "'"
Else ' not equals
  andWhere = " AND " & tablename & fieldname & " <> '" & field & "'"
End If
End Function

Private Sub whereField(fieldname As String, Optional tablename As String = "prop", _
                       Optional joinname As String, Optional andmore As String)
' add restrictions to SQL prop_id query FROM and WHERE clauses for field.
' Field values come from andWhere ~(in list, like, or =)
' field names: With one argument, field is in prop table, already present.
' With three, we need to add table to FROM, join with prop table, and restrict field
' For convenience/readability: If field, join start with _, prepend table name: e.g. prop_stts,_abbr -> prop_stts.prop_stts_abbr
Dim andclause, tablealias As String

If Left(fieldname, 1) = "_" Then fieldname = tablename & fieldname ' expand abbreviated names
If Left(joinname, 1) = "_" Then joinname = tablename & joinname
andclause = andWhere(tablename & ".", fieldname)
If Len(andclause) > 2 Then
    myPidsWhere = myPidsWhere & andclause & andmore & vbLf
    If tablename <> "prop" Then ' need to join a new table to prop
       If InStr(tablename, ".") = 0 Then tablename = "csd." & tablename 'fully qualify, if not already
       tablealias = Mid(tablename, InStrRev(tablename, ".") + 1)
       myPidsFrom = myPidsFrom & "JOIN " & tablename & " " & tablealias _
        & " ON prop." & joinname & " = " & tablealias & "." & joinname & vbLf
    End If
End If

End Sub


Sub ParseFields()
'identify the proposals, using the flexible subroutines above
    myPidsFrom = "FROM csd.prop prop" & vbNewLine
    myPidsWhere = ""
  With InputTab
    If hasValue("from_date") Then myPidsWhere = myPidsWhere & "AND prop.nsf_rcvd_date >= {ts '" & Format(.Range("from_date"), "yyyy-mm-dd hh:mm:ss") & "'} " & vbNewLine
    If hasValue("to_date") Then myPidsWhere = myPidsWhere & "AND prop.nsf_rcvd_date <= {ts '" & Format(.Range("to_date"), "yyyy-mm-dd hh:mm:ss") & "'} " & vbNewLine
    If hasValue("dd_from_date") Then myPidsWhere = myPidsWhere & "AND prop.dd_rcom_date >= {ts '" & Format(.Range("dd_from_date"), "yyyy-mm-dd hh:mm:ss") & "'} " & vbNewLine
    If hasValue("dd_to_date") Then myPidsWhere = myPidsWhere & "AND prop.dd_rcom_date <= {ts '" & Format(.Range("dd_to_date"), "yyyy-mm-dd hh:mm:ss") & "'} " & vbNewLine
  End With
    Call whereField("pgm_annc_id")
    Call whereField("org_code")
    Call whereField("pgm_ele_code")
    Call whereField("obj_clas_code")
    Call whereField("prop_titl_txt")
    Call whereField("pm_ibm_logn_id")
    Call whereField("pi_id")
    Call whereField("inst_id")
    Call whereField("_abbr", "prop_stts", "_code")
    Call whereField("_abbr", "natr_rqst", "_code")
    Call whereField("dir_div_abbr", "org", "_code")
    Call whereField("panl_id", "panl_prop", "prop_id")
    Call whereField("_code", "prop_atr", "prop_id", " AND prop_atr.prop_atr_type_code = 'PRC'")
    If hasValue("routing.CODE") Then
       myPidsFrom = myPidsFrom & "JOIN csd.prop_subm_ctl_vw psc ON prop.prop_id = psc.prop_id " & vbNewLine _
              & "JOIN csd.routing_vw routing ON psc.TEMP_PROP_ID = routing.TEMP_PROP_ID " & vbNewLine
       myPidsWhere = myPidsWhere & andWhere("", "routing.CODE") & vbNewLine
    End If
    If Len(myPidsWhere) < 3 Then
        myPidsFrom = "" ' retrieve do nothing if no restrictions are applied
    Else
        myPidsWhere = Replace(myPidsWhere, "AND", "WHERE", 1, 1) ' replace first AND with WHERE
    End If
End Sub


Function InputSQL() As String
' create the query string that begins all proposal-based queries
Dim addProps As String
Dim myPidsSelect As String

Call ParseFields
myPidsSelect = "SELECT CASE WHEN prop.lead_prop_id IS NULL " _
  & "THEN 'I' WHEN prop.lead_prop_id <> prop.prop_id THEN 'N' ELSE 'L' END AS ILN," & vbNewLine _
  & "isnull(prop.lead_prop_id,prop.prop_id) AS lead, prop.prop_id" & vbNewLine
  
addProps = IDsFromColumnRange(myPidsSelect _
        & "INTO #myPids FROM csd.prop prop WHERE prop.prop_id IN", InputTab.Range("add_prop_ids"))

If Len(addProps) < 1 Then
  If Len(myPidsFrom) < 1 Then
     MsgBox "You must either include restrictions (esp. dates) or include proposal numbers in the Inclusion table."
     End
  End If
  addProps = myPidsSelect & "INTO #myPids " ' need INTO
ElseIf Len(myPidsFrom) > 0 Then
  addProps = addProps & "UNION ALL " & myPidsSelect  ' need UNION
End If

'identify the proposals, using the flexible subroutines above
InputSQL = "SET NOCOUNT ON" & vbNewLine & addProps & myPidsFrom & myPidsWhere _
& "INSERT INTO #myPids SELECT CASE WHEN prop.lead_prop_id <> prop.prop_id THEN 'N' ELSE 'L' END as ILN, p.lead, prop.prop_id" & vbNewLine _
& "FROM #myPids p, csd.prop prop WHERE p.ILN <> 'I' AND p.lead = prop.lead_prop_id" & vbNewLine _
& "SELECT prop.nsf_rcvd_date, nullif(prop.dd_rcom_date,'1900-01-01') AS dd_rcom_date," & vbNewLine _
& "prop.pgm_annc_id, o2.dir_div_abbr as Dir, prop.org_code, CASE WHEN prop.org_code <> prop.orig_org_code THEN prop.orig_org_code END AS origORG, " & vbNewLine _
& "prop.pgm_ele_code+' - '+pgm_ele_name as Pgm, CASE WHEN prop.pgm_ele_code <> prop.orig_pgm_ele_code THEN prop.orig_pgm_ele_code ELSE ' ' END AS origPEC," & vbNewLine _
& "prop.pm_ibm_logn_id as PO, prop.obj_clas_code, natr_rqst.natr_rqst_abbr, prop_stts.prop_stts_abbr, p.ILN, p.lead, org.dir_div_abbr as Div, p.prop_id," & vbNewLine _
& "pi.pi_last_name, pi.pi_frst_name, inst.inst_shrt_name AS inst_name, inst.st_code, pi.pi_emai_addr," & vbNewLine _
& "prop.prop_titl_txt, prop.rqst_dol, prop.rqst_eff_date, prop.rqst_mnth_cnt, prop.cntx_stmt_id, prop.prop_stts_code, prop_stts.prop_stts_txt, prop.inst_id, prop.pi_id" & vbNewLine _
& "INTO #myProps FROM (SELECT DISTINCT * FROM #myPids " & IDsFromColumnRange("WHERE prop_id NOT IN", InputTab.Range("omit_prop_ids")) & ") p" & vbNewLine _
& "JOIN csd.prop prop ON p.prop_id = prop.prop_id" & vbNewLine _
& "JOIN csd.org org ON prop.org_code = org.org_code" & vbNewLine _
& "JOIN csd.org o2 ON left(prop.org_code,2)+'000000' = o2.org_code" & vbNewLine _
& "JOIN csd.pgm_ele pe ON prop.pgm_ele_code = pe.pgm_ele_code" & vbNewLine _
& "JOIN csd.inst inst ON prop.inst_id = inst.inst_id" & vbNewLine _
& "JOIN csd.pi_vw pi ON prop.pi_id = pi.pi_id" & vbNewLine _
& "JOIN csd.prop_stts prop_stts ON prop.prop_stts_code = prop_stts.prop_stts_code" & vbNewLine _
& "JOIN csd.natr_rqst natr_rqst ON natr_rqst.natr_rqst_code = prop.natr_rqst_code" & vbNewLine _
& "WHERE pi.prim_addr_flag='Y'" & vbNewLine _
& "ORDER BY lead, ILN DROP TABLE #myPids" & vbNewLine
End Function


Sub preQuery(mySheet)
' count the number of things that are to be retrieved by query
Call ParseFields

Dim QT As QueryTable
Dim i As Long
Dim errCount As Long
errCount = 0

    On Error GoTo cantRefresh ' if sheet or query doesn't exist, skip it.
    Set QT = mySheet.ListObjects("CountTable").QueryTable
retryQuery:
    With QT
        .CommandText = "SELECT count(*) as num_retrieved " & myPidsFrom & myPidsWhere
        .Refresh (False)
    End With
    DoEvents
cleanExit:
    On Error GoTo 0
    Exit Sub
   
cantRefresh:
    If Err.Number = 1004 Then 'perhaps odbc timeout
       errCount = errCount + 1
       If errCount < 2 Then Resume ' retry query once more in case we timed out
    End If
    MsgBox "Error " & Err.Number & ": Cannot count number retrieved on " & mySheet.name & " Aborting. " & Err.description
    End
End Sub








