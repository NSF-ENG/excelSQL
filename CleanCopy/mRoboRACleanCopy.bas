Attribute VB_Name = "mRoboRACleanCopy"
Option Explicit
' CleanCopy.v1.2.dotm        by Jack Snoeyink   Dec 15, 2017
' Macros and functions for a word template/add-in that can clean up special
' characters (left/right quotes, em&en dashes, ellipses, bullets, & common accented characters)
' that would otherwise turn to ? in the interactive panel system or eJacket.
' This is designed to be deployed as an addin in Word.
' Reference needed:  Word Object Library (works with >=15.0, perhaps earlier.)
' On a PC, place it in %appdata%/microsoft/word/startup (create if it doesn't exist)
' On a Mac, use Macros Add-in menu.
' Should be digitally signed, and not disabled by group policy.
' Can customize ribbon to add the Public macros
'
' Security did not want this deployed as an add-in, so it will be at a particular location in sharepoint
'
#If Mac Then
 #If MAC_OFFICE_VERSION >= 15 Then
  #If VBA7 Then ' 64-bit Excel 2016 for Mac
   Declare PtrSafe Function GetTickCount Lib _
"/Applications/Microsoft Excel.app/Contents/Frameworks/MicrosoftOffice.framework/MicrosoftOffice" () As Long
  #Else ' 32-bit Excel 2016 for Mac
   Declare Function GetTickCount Lib _
"/Applications/Microsoft Excel.app/Contents/Frameworks/MicrosoftOffice.framework/MicrosoftOffice" () As Long
  #End If
 #Else
  #If VBA7 Then ' does not exist, but why take a chance
   Declare PtrSafe Function GetTickCount Lib _
"Applications:Microsoft Office 2011:Office:MicrosoftOffice.framework:MicrosoftOffice" () As Long
  #Else ' 32-bit Excel 2011 for Mac
   Declare Function GetTickCount Lib _
"Applications:Microsoft Office 2011:Office:MicrosoftOffice.framework:MicrosoftOffice" () As Long
  #End If
 #End If
#Else 'PC
    #If VBA7 Then
      Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As LongPtr)
    #Else
      Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
    #End If
#End If

Private Sub BusyWait(t As Long)
#If Mac Then
' busy wait t clock ticks
    Dim EndTick As Long
    EndTick = GetTickCount + t
    Do
        DoEvents
    Loop Until GetTickCount >= EndTick
#Else 'PC
    DoEvents
    Sleep (t)
    DoEvents
#End If
End Sub

Private Function FixIPSText(s As String) As String
' Replace special characters in String s with ascii equivalents and return
' Each replacement is tried on entire string, so time is (#possible replacements)x(string length).
' The special characters are in the text of thisDocument, namely the template.

Dim repl() As String
Dim i As Long
repl = Split(Replace(ThisDocument.Content.Text, vbLf, vbCr), vbCr) ' handle cr/lf
'Debug.Print LBound(repl) & ", " & UBound(repl)
    For i = LBound(repl) + 1 To UBound(repl)
       If Len(repl(i)) > 1 Then
         'Debug.Print Left(repl(i), 1) & AscW(Left(repl(i), 1)) & " ->" & Mid(repl(i), 3) & "<"
          s = Replace(s, Left(repl(i), 1), Mid(repl(i), 3))
       End If
    Next i
FixIPSText = Replace(s, Chr(11), Chr(13)) ' eJ doesn't like vertical tab
End Function

Private Function StripDoubleBrackets(s As String) As String
' does not handle nesting.
Dim i As Long, j As Long, k As Long, lenS As Long

i = InStrRev(s, "[[")
j = InStrRev(s, "]]")
If j < i Then MsgBox "Warning: last open bracket has no close." & vbNewLine & Right(s, Len(s) - i + 1)
Do While j > 0
    i = InStrRev(s, "[[", j)
    k = InStrRev(s, "]]", j - 1)
    If i < k Then
      MsgBox "Warning: Consecutive close comment brackets with no open." & vbNewLine & Mid(s, k, j - k + 2)
      j = k
    End If
    If i < 1 Then
      MsgBox "Warning: missing open comment brackets for first close." & vbNewLine & Left(s, j + 1)
      j = 0
    Else
      s = Left(s, i - 1) & Right(s, Len(s) - j - 1)
      j = InStrRev(s, "]]", i)
    End If
Loop

StripDoubleBrackets = s
End Function
Private Sub CopyText(Text As String)
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

Public Sub CleanIPSCopy()
' Clean up special characters in selection and copy to clipboard.
' If nothing is selected, clean and copy the entire document
' Suggested name CleanCopy if you add to Ribbon
If Selection.Characters.Count < 2 Then Selection.WholeStory
 CopyText (FixIPSText(Selection.Text))
 Selection.Collapse
End Sub

Public Sub StripCleanIPSCopy()
' Strip [[instructions in double square brackets]], clean up special characters in selection
' and copy to clipboard. If nothing is selected, strip, clean, and copy entire document
' Suggested name [CleanCopy] if you add to Ribbon
 If Selection.Characters.Count < 2 Then Selection.WholeStory
 CopyText (FixIPSText(StripDoubleBrackets(Selection.Text)))
 Selection.Collapse
End Sub

Private Sub VisitEJacket(prop_id As String, EJlastSection As String)
'Try to visit the appropriate eJ section for the given prop_id to get ready to paste
'If eJ hyperlink structure changes, most changes are here, but consider also the Clean...2EJ functions below,
'which provide the EJlastSection string to choose to add Abstract, PO comment, or Review Analysis.
'Assumse prop_id is well formatted (currently as a string of 7 digits.)
    Dim i As Long
    
    On Error GoTo ErrHandler
    With ActiveDocument
    .FollowHyperlink ("https://www.ejacket.nsf.gov/ej/showProposal.do?Continue=Y&ID=" & prop_id)
    BusyWait (1000)
    .FollowHyperlink ("https://www.ejacket.nsf.gov/ej/" & EJlastSection)
    DoEvents
    End With
ExitHandler:
    Exit Sub
ErrHandler:
    If Err.Number = 4198 Then
        MsgBox ("Unable to visit eJacket hyperlinks; Please visit EJ for proposal " & prop_id & " to paste in appropriate document.")
    Else
        MsgBox ("Unexpected error " & Err.Number & ", " & Err.Description & vbNewLine & "Please visit EJ for proposal " & prop_id & " to paste in appropriate document.")
    End If
    Resume ExitHandler:
End Sub

Private Function InputPropid(docName As String) As String
  ' check if the document starts with a [[prop_id comment.  If not,
  ' ask user for prop_id; returns either a string of seven digits or empty string; warns if malformed prop_id entered.
  
  Dim prop_id As String
  Dim i As Integer, j As Integer
  
  On Error Resume Next
  prop_id = Left(ActiveDocument.Content.Text, 16) ' get initial characters, then look for consecutive digits
  If Err.Number = 0 Then
     i = 1
     While i < 16 And (Mid(prop_id, i, 1) < "0" Or Mid(prop_id, i, 1) > "9")
       i = i + 1
     Wend
     j = i + 1
     While j < 16 And (Mid(prop_id, j, 1) >= "0" And Mid(prop_id, j, 1) <= "9")
       j = j + 1
     Wend
     prop_id = Mid(prop_id, i, j - i)
   End If
   If (Len(prop_id) <> 7 Or Val(prop_id) = 0) Then
      prop_id = VBA.Format$(VBA.Trim$(InputBox("7 digit proposal id for this " & docName, "Enter prop_id")), "0000000")
  If (Len(prop_id) <> 7 Or Val(prop_id) = 0) Then
    If Len(prop_id) > 0 Then MsgBox ("Did not get a valid prop_id " & prop_id)
    prop_id = ""
  End If
  InputPropid = prop_id
End Function

Public Sub Abst2EJ()
' Asks for a proposal id, which can be in clipboard, cleans the selection to clipboard,
' and visits the abstract page in eJacket, ready to paste.
  Dim prop_id As String
  prop_id = InputPropid("Project Abstract")
  If Not prop_id = "" Then
    Call CleanIPSCopy
    'https://www.ejacket.nsf.gov/ej/processProposalAbstract.do?dispatch=showAdd
    Call VisitEJacket(prop_id, "processProposalAbstract.do?dispatch=showAdd")
  End If
End Sub

Public Sub POCom2EJ()
' Asks for a proposal id, which can be in clipboard, strips [[comments]] and cleans selection
' to clipboard, and visits the PO Comment page in eJacket, ready to paste.
    Dim prop_id As String
    prop_id = InputPropid("PO comment")
    If Not prop_id = "" Then
      Call StripCleanIPSCopy ' strip [[comments]], too
      'https://www.ejacket.nsf.gov/ej/processPoComment.do?dispatch=showAdd
      Call VisitEJacket(prop_id, "processPoComment.do?dispatch=showAdd")
    End If
End Sub

Public Sub RA2EJ()
' Asks for a proposal id, which can be in clipboard, strips [[comments]] and cleans selection
' to clipboard, and visits the Review Analysis in eJacket, ready to paste.
    Dim prop_id As String
    prop_id = InputPropid("Review Analysis")
    If Not prop_id = "" Then
      Call StripCleanIPSCopy ' strip [[comments]], too
      'https://www.ejacket.nsf.gov/ej/processReviewAnalysis.do?dispatch=add&uniqId=1749173jsnoeyin
      Call VisitEJacket(prop_id, "processReviewAnalysis.do?dispatch=add&uniqId=" & prop_id & LCase(Left(Environ("USERNAME"), 7)))
    End If
End Sub

Sub RoboRAStripCleanCopy()
' Don't add to the Ribbon: This is for RoboRA to call in RA templates on either Mac or PC versions of Word.
' Assumes that this macro is called from a field code, and that it is followed by a prop_id in a private field.
' Strips and cleans whole document, and opens eJ Review Analysis, ready to paste.
    Dim prop_id As String
    prop_id = VBA.Format$(Mid$(Selection.Fields(2).Code, 10, 7), "0000000") ' pull prop_id from private field
    If Val(prop_id) = 0 Then MsgBox ("You may be running the macro in the Template rather than a merge document because the prop_id is " & prop_id)
    Selection.Collapse
    
    Call StripCleanIPSCopy
    
    With ActiveDocument ' save RA with ReadOnlyRecommend to indicate it has been uploaded
        .ReadOnlyRecommended = True
        .Save
    End With
    'https://www.ejacket.nsf.gov/ej/processReviewAnalysis.do?dispatch=add&uniqId=1749173jsnoeyin
    Call VisitEJacket(prop_id, "processReviewAnalysis.do?dispatch=add&uniqId=" & prop_id & LCase(Left(Environ("USERNAME"), 7)))
End Sub






