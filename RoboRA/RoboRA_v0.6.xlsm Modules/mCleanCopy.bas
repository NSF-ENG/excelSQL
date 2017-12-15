Attribute VB_Name = "mCleanCopy"
Option Explicit
Function FixIPSText(s As String) As String
'
' Replace special characters
'

Dim fromC As String, toC As String
Dim i As Long
fromC = ChrW(160) & "–”“’‘•ãèéàáâåçêëìíîïòóôõùúû"
toC = " -""""''*aeeaaaaceeiiiioooouuu"

'MsgBox Len(fromC) & " : " & Len(toC)
    s = Replace(Replace(Replace(Replace(Replace(s, "…", "..."), "—", "--"), "ä", "ae"), "ö", "oe"), "ü", "ue") ' mulitcharacter replaces —äöü

    For i = 1 To Len(fromC)
       'MsgBox AscW(Mid(fromC, i, 1)) & " - " & AscW(Mid(toC, i, 1))
       s = Replace(s, Mid(fromC, i, 1), Mid(toC, i, 1))
    Next i
    For i = 1 To Len(s) ' replace anything I missed with ?
      If AscW(Mid(s, i, 1)) > 127 Then s = Replace(s, Mid(s, i, 1), "?")
    Next i
FixIPSText = s
End Function

Function StripDoubleBrackets(s As String) As String
' does not handle nesting.
Dim i As Long, j As Long, k As Long, lenS As Long

j = InStrRev(s, "]]")
Do While j > 0
    i = InStrRev(s, "[[", j)
    k = InStrRev(s, "]]", j - 1)
    If i < k Then
      MsgBox "Warning: Consecutive close comment brackets with no open." & vbNewLine & Mid(s, k, j - k + 1)
      j = k
    End If
    If i < 1 Then
      MsgBox "Warning: missing open comment brackets for first close." & vbNewLine & Left(s, j)
    Else
      s = Left(s, i - 1) & Right(s, Len(s) - j - 1)
    End If
    j = InStrRev(s, "]]", i)
Loop

StripDoubleBrackets = s
End Function

