Attribute VB_Name = "mSQLcode"
Option Explicit
Sub saveSQLcode()
' assume text in clipboard contains substrings like --[name SQL code --]name
' and defines or updates Range(name) with the SQL code
' in the bottom of cols A:C in HiddenSettings
   'VBA Macro using late binding to edit text in clipboard.
    'Modified by Jack Snoeyink Aug 2016 from original by Justin Kay, 8/15/2014
    'Thanks to http://akihitoyamashiro.com/en/VBA/LateBindingDataObject.htm
    Dim cb As Object
    Dim a() As String
    Dim s As String, varname As String
    Dim i As Long, j As Long, k As Long, r As Long
    Dim rng As Range

    #If Mac Then
        Set cb = New DataObject
    #Else
        Set cb = CreateObject("new:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")
    #End If
    cb.GetFromClipboard
    
'    If cb.ContainsText Then
    s = cb.GetText()
    i = VBA.InStr(s, "--[")
    If i = 0 Then MsgBox (" No --[rangename --]rangename  pairs found in text " & vbNewLine & VBA.Left$(s, 100))
    While i > 0
        j = VBA.InStr(i, s, vbLf)
        If j = 0 Then
            MsgBox ("Need line feed (vbLF) here: " & VBA.Mid$(s, i, 100))
            End
        End If
        varname = VBA.Mid$(s, i + 3, j - i - 4)
        On Error Resume Next
        Set rng = HiddenSettings.Range(varname) ' check if varname is defined
        If Err.Number = 0 Then
         r = rng.Row ' where this varname lives now
        ElseIf Err.Number = 1004 Then
         r = HiddenSettings.Range("C" & HiddenSettings.Rows.count).End(xlUp).Row + 1 ' one past last row
         HiddenSettings.Names.Add Name:=varname, RefersTo:="=" & HiddenSettings.Name & "!$C$" & r
        Else
         MsgBox ("Unexpected error " & Err.Number & ":" & Err.Description)
         End
        End If
        On Error GoTo 0
        k = VBA.InStr(j, s, "--]" & varname)
        If k = 0 Then
           k = VBA.InStr(j, s, "--]")
           If k = 0 Then
            MsgBox ("Error: unterminated --[" & varname)
            k = Len(s) + 1
           Else
            MsgBox ("Warning: terminating --[" & varname & " with " & VBA.Mid$(s, k, 15))
           End If
        End If
        HiddenSettings.Cells(r, 1).Value = varname
        HiddenSettings.Cells(r, 2).Value = Now
        HiddenSettings.Cells(r, 3).Value = VBA.Mid$(s, j + 1, k - j - 1)
        i = VBA.InStr(k, s, "--[")
    Wend
    cb.Clear
    Set cb = Nothing
End Sub
