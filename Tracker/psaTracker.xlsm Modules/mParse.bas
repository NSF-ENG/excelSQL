Attribute VB_Name = "mParse"
Option Explicit

Global mySQLWhere As String
Global mySQLJoins As String

Function hasValue(rangeName As String) As Boolean
' check if there is non-example data in this range
 With ActiveSheet.Range(rangeName)
   hasValue = (Trim$(.Value) <> "" And Left$(.Value, 3) <> "eg:")
  End With
End Function

Function andWhere(tablename As String, fieldname As String, Optional notPreamble As String = " NOT (", Optional andMore As String = "")
' Flexible WHERE clause construction for SQL queries with text fields.
' Warning: trims spaces in value;
' Field values have 5 cases:
'       blank or eg: parameter: nothing,
'       comma-separated list:   AND field IN ('val1','val2',...,'valN'),
'       val1::val2:             AND field BETWEEN 'val1' AND 'val2'
'       has SQL wildcard        AND field LIKE value,
'       any other value         AND field = value
'    First character ~ negates any of the above, by applying " NOT (" or a more complex notPreamble.
'    andMore adds other restrictions
' E.g., for PRCs, which are many2many, ~ as NOT ( will mean that the proposal has to have some PRC that is not in the list.
' We can send notPreamble="NOT EXISTS (SELECT * FROM csd.prop_atr pa WHERE pa.prop_id=prop.prop_id " and
' andMore="AND pa.prop_atr_type_code='PRC'" ' to force a "not any", meaning it cannot have any PRC in the list.


Dim field As String
Dim optNeg As String
Dim hasComma As Boolean, hasRange As Boolean, hasSqlWildcard As Boolean
optNeg = "("
andWhere = ""
On Error Resume Next
field = Trim(ActiveSheet.Range(fieldname).Value) ' Warning: trims spaces on values

On Error GoTo 0
If Err.Number > 0 Then
    If Err.Number = 1004 Then
      If MsgBox("Named range " & fieldname & " may have been lost from " & ActiveSheet.Name & vbNewLine & _
         "Continuing without, but if a cell was moved by Copy/Paste consider Undo and Paste Values Only (mac:Paste Special>Text)", vbOKCancel) <> vbOK Then End
    Else
      If MsgBox("Unexpected error: " & Err.Number & ":" & Err.Description & ". Continuing", vbOKCancel) <> vbOK Then End
    End If
End If
If Left(field, 3) = "eg:" Then field = "" ' ignore example fields

If Left(field, 1) = "~" Then ' have negation
  optNeg = notPreamble ' default should be NOT (, but can be NOT EXIST (SELECT... or NOT IN (SELECT...
  field = Right(field, Len(field) - 1)
End If

hasComma = (InStr(field, ",") > 0) 'have list
hasRange = (InStr(field, "::") > 0) 'have range
hasSqlWildcard = (InStr(field, "%") > 0) Or (InStr(field, "_") > 0) Or ((InStr(field, "[") > 0) And (InStr(field, "]") > 0))
If (hasSqlWildcard And (hasRange Or hasComma)) Then
 MsgBox "Can't use SQL wildcards in a range or comma separated list: " & fieldname & " = " & field
 End
End If
If Len(tablename) > 1 Then ' make sure we end in "."
    If Right(tablename, 1) <> "." Then tablename = tablename & "."
End If

If Len(field) < 1 Then ' do nothing; andwhere is blank
ElseIf hasComma Then ' IN/NOT IN list
  andWhere = " AND " & optNeg & tablename & fieldname & " IN ('" & Replace(Replace(Join(Split(Replace(Replace(field, """", ""), "'", ""), ","), "','") & "') ", " '", "'"), "' ", "'") & andMore & ")"
ElseIf hasRange Then ' BETWEEN / NOT BETWEEN
  andWhere = " AND " & optNeg & tablename & fieldname & " BETWEEN '" & Replace(Replace(Left(field, InStr(field, "::") - 1), """", ""), "'", "") & "' AND '" & Replace(Replace(Mid(field, InStr(field, "::") + 2), """", ""), "'", "") & "' " & andMore & ")"
ElseIf hasSqlWildcard Then ' LIKE / NOT LIKE
  andWhere = " AND " & optNeg & tablename & fieldname & " LIKE '" & Replace(Replace(field, """", ""), "'", "") & "' " & andMore & ")"
Else ' = / NOT =
  andWhere = " AND " & optNeg & tablename & fieldname & " = '" & Replace(Replace(field, """", ""), "'", "") & "' " & andMore & ")"
End If
End Function

Sub whereField(fieldname As String, Optional tablename As String = "prop", Optional joinname As String, _
                       Optional isIntField As Boolean = False, Optional notPreamble As String = " NOT (", Optional andMore As String = "")
' add restrictions to SQL prop_id query FROM and WHERE clauses for field.
' using two global variables, mySQLWhere and mySQLJoins
'
' Field values come from andWhere ~(in list, between, like, or =); see that Function on notPreamble and andMore
' field names: With one argument, field is in prop table, already present.
' With three, we need to add table to FROM, join with prop table, and restrict field
' For convenience/readability: If field, join start with _, prepend table name: e.g. prop_stts,_abbr -> prop_stts.prop_stts_abbr
' Most fields are strings, so are quoted: IsIntField:=True will strip quotes to allow integer parameters.
Dim andclause As String
Dim tablealias As String

If Left(fieldname, 1) = "_" Then fieldname = tablename & fieldname ' expand abbreviated names
If Left(joinname, 1) = "_" Then joinname = tablename & joinname
andclause = andWhere(tablename & ".", fieldname, notPreamble, andMore)
If Len(andclause) > 2 Then
    If isIntField Then andclause = Replace(andclause, "'", "") ' for integer field, strip quotes.
    mySQLWhere = mySQLWhere & andclause & andMore & vbLf
    If tablename <> "prop" Then ' need to join a new table to prop
       If InStr(tablename, ".") = 0 Then tablename = "csd." & tablename 'fully qualify, if not already
       tablealias = Mid(tablename, InStrRev(tablename, ".") + 1)
       mySQLJoins = mySQLJoins & "JOIN " & tablename & " " & tablealias _
        & " ON prop." & joinname & " = " & tablealias & "." & joinname & vbLf
    End If
End If
End Sub

Function IDsFromColumnRange(prefix As String, tbl As ListObject) As String
Dim ids As String
' make comma separated list of column ids, with sql prefix
IDsFromColumnRange = ""
If tbl Is Nothing Then
   MsgBox ("Error: can't find table on " & ActiveSheet.Name & " for " & prefix & vbNewLine & "This is a bug in the VBA code, or the table was deleted. Ignoring & continuing.")
   Exit Function
End If
If tbl.DataBodyRange Is Nothing Then Exit Function
With tbl.DataBodyRange
    If .Rows.Count < 2 Then
        ids = .Value
    Else ' we have at least two, and can use transpose
        ids = Join(Application.Transpose(.Value), "','") ' quote column entries and make a comma-separated row
    End If
    ids = "'" & Replace(Replace(ids, " ", ""), Chr(160), "") & "'" ' strip spaces (visible and invisible) and ,'' from string.
    ids = Replace(ids, ",''", "")  ' strip blank column entries
    'MsgBox "please check your ids : " + ids
    If Len(ids) > 2 Then IDsFromColumnRange = prefix & " (" & ids & ")" & vbLf
End With
End Function

Public Function tableFromRange(rng As String, types As String, tableDefn As String, Optional tempTable As String = "#myTable")
'Build SQL to create table tempTable with tableDefn and populate columns from range on the ActiveSheet.
'Rows where the first column is blank are omitted; all other blanks become NULL.
'Returns empty string if first column is all blank.
'string types is i for integer columns; all others are quoted.
'Here are the results of an example, assuming the ActiveSheet has this table populated...
'debug.Print tableFromRange("IncludeTableRecent","cic", "(prop_id char(7) NOT NULL, budg_yr int NULL, split_id char(2) NULL) ")
' INSERT INTO #myTable SELECT 'a',1,'3'
' UNION ALL SELECT 'b',2,NULL
' UNION ALL SELECT 'c',NULL,NULL
' UNION ALL SELECT 'aa',11,'33'
' UNION ALL SELECT 'bb',22,NULL
' UNION ALL SELECT 'cc',NULL,NULL
Dim a As Variant
Dim i As Integer
Dim j As Integer
Dim s As String
Dim prefix As String
s = ""

a = ActiveSheet.Range(rng)
prefix = "CREATE TABLE " & tempTable & tableDefn & vbNewLine _
 & "INSERT INTO " & tempTable & " SELECT "

For i = LBound(a) To UBound(a)
  If a(i, 1).Value <> "" Then ' skip any row with blank first column
   If Left(types, 1) <> i Then
     s = s & prefix & "'" & a(i, 1).Value & "'" ' string
   Else
     s = s & prefix & a(i, 1).Value 'integer
   End If
   For j = 2 To UBound(a, 2) ' handle remaining columns
    If a(i, j).Value = "" Then
      s = s & ",NULL" 'null
    ElseIf Mid(types, j, 1) <> "i" Then
      s = s & ",'" & a(i, j).Value & "'" 'string
    Else
      s = s & "," & a(i, j).Value ' integer
    End If
   Next j
   prefix = vbNewLine & "UNION ALL SELECT "
  End If
Next i
tableFromRange = s
End Function



Private Sub TestParse()
  ActiveSheet.Range("pa.prop_atr_code").Value = "1234"
  Debug.Print andWhere("", "pa.prop_atr_code")
  ActiveSheet.Range("pa.prop_atr_code").Value = "'1234',""2222"",3333"
  Debug.Print andWhere("", "pa.prop_atr_code")
  ActiveSheet.Range("pa.prop_atr_code").Value = "'12%'"
  Debug.Print andWhere("", "pa.prop_atr_code")
  ActiveSheet.Range("pa.prop_atr_code").Value = "1234::2222"
  Debug.Print andWhere("", "pa.prop_atr_code")
  ActiveSheet.Range("pa.prop_atr_code").Value = "~'1234'"
  Debug.Print andWhere("", "pa.prop_atr_code")
  ActiveSheet.Range("pa.prop_atr_code").Value = "~'1234',""2222"",3333"
  Debug.Print andWhere("", "pa.prop_atr_code")
  ActiveSheet.Range("pa.prop_atr_code").Value = "~'12%'"
  Debug.Print andWhere("", "pa.prop_atr_code")
  ActiveSheet.Range("pa.prop_atr_code").Value = "~1234::2222"
  Debug.Print andWhere("", "pa.prop_atr_code")
  ActiveSheet.Range("pa.prop_atr_code").Value = "1234::2222"
  Debug.Print andWhere("", "pa.prop_atr_code", "NOT EXISTS (SELECT * FROM csd.prop_atr pa WHERE pa.prop_id=prop.prop_id AND ", "AND pa.prop_atr_type_code='PRC'")
  ActiveSheet.Range("pa.prop_atr_code").Value = "~1234::2222"
  Debug.Print andWhere("", "pa.prop_atr_code", "NOT EXISTS (SELECT * FROM csd.prop_atr pa WHERE pa.prop_id=prop.prop_id AND ", "AND pa.prop_atr_type_code='PRC'")

  ActiveSheet.Range("pa.prop_atr_code").Value = "'12%',2222,3333"
  'Debug.Print andWhere("", "pa.prop_atr_code") ' error
End Sub
