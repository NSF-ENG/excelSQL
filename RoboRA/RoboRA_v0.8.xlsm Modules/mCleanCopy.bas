Attribute VB_Name = "mCleanCopy"
Option Explicit
Function FixIPSText(s As String) As String
'
' Replace special characters
'

Dim fromC As String, toC As String
Dim i As Long
'Do not edit & save on Mac or it well munge the special characters
fromC = ChrW(160) & "–”“’‘•ãèéàáâåçêëìíîïòóôõùúû"
toC = " -""""''*aeeaaaaceeiiiioooouuu"

'MsgBox Len(fromC) & " : " & Len(toC)
    s = Replace(Replace(Replace(Replace(Replace(s, "…", "..."), "—", "--"), "ä", "ae"), "ö", "oe"), "ü", "ue") ' mulitcharacter replaces —äöü

    For i = 1 To Len(fromC)
       'MsgBox AscW(Mid(fromC, i, 1)) & " - " & AscW(Mid(toC, i, 1))
       s = Replace(s, VBA.Mid$(fromC, i, 1), VBA.Mid$(toC, i, 1))
    Next i
    For i = 1 To Len(s) ' replace anything I missed with ?
      If AscW(Mid(s, i, 1)) > 127 Then s = Replace(s, VBA.Mid$(s, i, 1), "?")
    Next i
FixIPSText = Replace(s, Chr(11), Chr(13)) ' fix vertical tab to CR
End Function

Function StripDoubleBrackets(s As String) As String
' does not handle nesting.
Dim i As Long, j As Long, k As Long, lenS As Long

j = InStrRev(s, "]]")
Do While j > 0
    i = InStrRev(s, "[[", j)
    k = InStrRev(s, "]]", j - 1)
    If i < k Then
      MsgBox "Warning: Consecutive close comment brackets with no open." & vbNewLine & VBA.Mid$(s, k, j - k + 1)
      j = k
    End If
    If i < 1 Then
      MsgBox "Warning: missing open comment brackets for first close." & vbNewLine & VBA.Left$(s, j)
    Else
      s = VBA.Left$(s, i - 1) & VBA.Right$(s, Len(s) - j - 1)
    End If
    j = InStrRev(s, "]]", i)
Loop

StripDoubleBrackets = s
End Function

Sub CopyText(Text As String)
    'VBA Macro using late binding to copy text to clipboard.
    'By Justin Kay, 8/15/2014
    'Thanks to http://akihitoyamashiro.com/en/VBA/LateBindingDataObject.htm
    'Needs reference MS Office Object Library
    Dim cb As Object
    
    #If Mac Then
        Set cb = New DataObject
    #Else
        Set cb = CreateObject("new:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")
    #End If
    cb.Clear
    cb.SetText Text
    cb.PutInClipboard
    Set cb = Nothing
End Sub
