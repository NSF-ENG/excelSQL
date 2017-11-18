Attribute VB_Name = "PRCSIncludeExclude"
Public Function includePRCS(tablename As String, alias As String, fieldname As String, Optional andMore As String) As String 'used to include
'Variables
Dim field As String
Dim hasComma As Boolean
Dim hasSqlWildcard As Boolean

field = Trim(ActiveSheet.Range(fieldname).Value)
hasComma = (InStr(field, ",") > 0)
hasSqlWildcard = (InStr(field, "%") > 0) Or (InStr(field, "_") > 0) Or ((InStr(field, "[") > 0) And (InStr(field, "]") > 0))

If Left(field, 1) = "~" Then ' have negation
    field = Right(field, Len(field) - 1)
End If

If Len(alias) > 0 Then 'sometimes fields come with alias and period, if not, add it
    alias = alias & "."
End If

If (hasComma And hasSqlWildcard) Then ' cannot have commas and wildcards
    MsgBox "Can't have sql wildcards in a comma separated list: " & field
    End
End If
 
 

    If Len(field) < 1 Then 'nothing
      includePRCS = ""
      
    ElseIf hasComma = True Then 'e.g. and b.prop_id in ('12313123','23123435','75675688')
        includePRCS = " AND exists (SELECT prop_id from " & tablename & "where " & alias & fieldname & " IN ('" & Join(Split(Replace(Replace(field, """", ""), "'", ""), ","), "','") & "')" & andMore & ")"
  
    ElseIf hasSqlWildcard Then 'e.g. and p.prop_id like '%10202020%'

        includePRCS = " AND exists (SELECT prop_id from " & tablename & "where " & alias & fieldname & " LIKE '" & Replace(Replace(field, """", ""), "'", "") & "'" & andMore & ")"
  
    Else
        includePRCS = " AND exists (SELECT prop_id from " & tablename & "where " & alias & fieldname & " = '" & field & "'" & andMore & ")"

    End If

End Function
Public Function excludePRCS(tablename As String, alias As String, fieldname As String, Optional andMore As String) As String 'used to exclude
'Variables
Dim field As String
Dim hasComma As Boolean
Dim hasSqlWildcard As Boolean

field = Trim(ActiveSheet.Range(fieldname).Value) 'get field value
hasComma = (InStr(field, ",") > 0)
hasSqlWildcard = (InStr(field, "%") > 0) Or (InStr(field, "_") > 0) Or ((InStr(field, "[") > 0) And (InStr(field, "]") > 0))

If Len(alias) > 0 Then 'sometimes fields come with alias and period, if not, add it
    alias = alias & "."
End If

If Left(field, 1) = "~" Then ' have negation
    field = Right(field, Len(field) - 1)
End If

If (hasComma And hasSqlWildcard) Then ' cannot have commas and wildcards
    MsgBox "Can't have sql wildcards in a comma separated list: " & field
    End
End If

    If Len(field) < 1 Then 'nothing
      excludePRCS = ""
      
    ElseIf hasComma = True Then 'e.g. and b.prop_id in ('12313123','23123435','75675688')
    
        excludePRCS = " AND not exists (SELECT prop_id from " & tablename & "where " & alias & fieldname & " IN ('" & Join(Split(Replace(Replace(field, """", ""), "'", ""), ","), "','") & "')" & andMore & ")"
    
    ElseIf hasSqlWildcard Then 'e.g. and b.prop_id like '%10202020'

        excludePRCS = " AND not exists (SELECT prop_id from " & tablename & "where " & alias & fieldname & " LIKE '" & Replace(Replace(field, """", ""), "'", "") & "'" & andMore & ")"
  
    Else
        excludePRCS = " AND not exists (SELECT prop_id from " & tablename & "where " & alias & fieldname & " = '" & field & "'" & andMore & ")"

    End If

End Function
