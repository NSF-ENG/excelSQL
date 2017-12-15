Attribute VB_Name = "mCleanCopy"
Function FixIPSText(s As String) As String
' Replace special characters
' IMPORTANT: don't save this on a Mac; you will lose the special characters
' since the Mac VBE will convert them, but the PC won't convert them back.

' Edit this code on a PC only -- DON'T SAVE ON MAC
Dim fromC As String, toC As String
Dim i As Long
 toC = " -""""''*aaaaaAceeeeiiiioooooOuuu"
 fromC = ChrW(160) & "–”“’‘•ãàáâåÅçèéêëìíîïòóôõøØùúû"
 s = Replace(Replace(Replace(Replace(Replace(s, "…", "..."), "—", "--"), "ä", "ae"), "ö", "oe"), "ü", "ue") ' mulitcharacter replaces
'MsgBox Len(fromC) & " : " & Len(toC)
    For i = 1 To Len(fromC)
       'MsgBox AscW(Mid(fromC, i, 1)) & " - " & AscW(Mid(toC, i, 1))
       s = Replace(s, Mid(fromC, i, 1), Mid(toC, i, 1))
    Next i
    For i = 1 To Len(s) ' any I missed become ?, otherwise Macs will fail to paste
      If AscW(Mid(s, i, 1)) > 127 Then
      'Debug.Print (i & ">" & Mid(s, i, 1) & "<-" & AscW(Mid(s, i, 1)))
       s = Replace(s, Mid(s, i, 1), "?")
     End If
    Next i
FixIPSText = s
End Function

Private Sub CopyText(Text As String)
    'VBA Macro using late binding to copy text to clipboard.
    'By Justin Kay, 8/15/2014
    'Thanks to http://akihitoyamashiro.com/en/VBA/LateBindingDataObject.htm
    Dim MSForms_DataObject As Object
    
    #If Mac Then
        Set MSForms_DataObject = New DataObject
    #Else
        Set MSForms_DataObject = CreateObject("new:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")
     #End If
    MSForms_DataObject.SetText Text
    MSForms_DataObject.PutInClipboard
    Set MSForms_DataObject = Nothing
End Sub

Sub CopyStrippedIPSText()
' copy text, strip comments, and fix special characters disliked by IPS

    Selection.WholeStory
    CopyText (FixIPSText(StripDoubleBrackets(Selection.Text)))
    
End Sub

Sub CopyIPSText()
' copy text and fix special characters disliked by IPS

    Selection.WholeStory
    CopyText (FixIPSText(Selection.Text))
End Sub

Function StripDoubleBrackets(s As String) As String
' does not handle nesting.
Dim i, j, k, lenS As Long

j = InStrRev(s, "]]")
Do While j > 0
    i = InStrRev(s, "[[", j)
    k = InStrRev(s, "]]", j - 1)
    If i < k Then
      MsgBox "Warning: Consecutive close comment brackets with no open." & vbNewLine & Mid(s, k, j - k + 1)
      j = k
    End If
    If i < 1 Then
      MsgBox "Warning: Consecutive close comment brackets with no open." & vbNewLine & Left(s, j)
    Else
      s = Left(s, i - 1) & Right(s, Len(s) - j - 1)
    End If
    j = InStrRev(s, "]]", i)
Loop

StripDoubleBrackets = s
End Function



