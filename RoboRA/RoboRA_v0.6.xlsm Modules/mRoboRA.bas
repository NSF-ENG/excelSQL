Attribute VB_Name = "mRoboRA"
Option Explicit
' Utility routines specific to RoboRA
Public Function SummarizeQuestMarks(abstr As String) As String
' Summarize ? from string that may have been converted from quotes, dashes, or other special characters.
Dim i As Long
Dim s As String
s = " "
i = InStrRev(abstr, "?")
Do While i > 0
  If i < 4 Then
    s = Mid(abstr, i, 8) & "|" & s
    Exit Do
  ElseIf Not Mid(abstr, i - 1, 3) Like "[a-zA-Z][?] " Then
    i = i - 3
    s = Mid(abstr, i, 8) & "|" & s
  End If
  i = InStrRev(abstr, "?", i - 1)
Loop
SummarizeQuestMarks = s
End Function

Private Sub test_SummarizeQuestMarks()
Debug.Print SummarizeQuestMarks("?Testing.? Is this OK? And this ? should not be??  ?Done?.")
End Sub

Function autoPasteRA(IE As InternetExplorerMedium, prop_id As String, RA As String) As String
' stuff RA into text box using mAutocoder functions
Dim i As Integer, j As Integer
Dim overwriteQ As Variant
Dim warn As String
warn = ""
overwriteQ = RoboRA.Range("overwrite_option").Value

If (Len(prop_id) = 7) Then ' Probably have a prop_id; go to Jacket
    IE.Navigate ("https://www.ejacket.nsf.gov/ej/showProposal.do?Continue=Y&ID=" & prop_id)
    Call myWait(IE)

    IE.Navigate ("https://www.ejacket.nsf.gov/ej/processReviewAnalysis.do?dispatch=add&uniqId=" & prop_id & LCase(Left(Environ("USERNAME"), 7)))
    Call myWait(IE)
    
    If IE.Document.getElementsByName("text")(0) Is Nothing Then
      warn = prop_id & " can't visit eJ RA" & vbNewLine
    Else
    With IE.Document.getElementsByName("text")(0)
      .Focus
      If (Len(.Value) < 10) Or (overwriteQ = 2) Then
       .Focus
       .Value = RA
      ElseIf (overwriteQ = 1) Then
        AppActivate Application.Caption
        DoEvents
        If (MsgBox("OK to overwrite existing RA for " & prop_id & vbNewLine & .Value, vbOKCancel) = vbOK) Then
         .Focus
         .Value = RA
        Else
          warn = prop_id & "Not overwritten." & vbNewLine
        End If
      End If
    End With
    Call myWait(IE)
    If Not IE.Document.getElementsByName("save")(0) Is Nothing Then
      IE.Document.getElementsByName("save")(0).Click
      Call myWait(IE)
    Else
      warn = prop_id & " can't save eJ RA" & vbNewLine
    End If
    End If
Else
  warn = prop_id & " not a prop_id" & vbNewLine
End If
autoPasteRA = warn
End Function


Sub List_Templates() ' list RA templates available (used by data validation)
Dim templateName As String
Dim nTemplates As Integer
nTemplates = 0
Application.ScreenUpdating = False
On Error GoTo ErrHandler
With Advanced.ListObjects("AvailableTemplates")
  If Not .DataBodyRange Is Nothing Then .DataBodyRange.Delete
  templateName$ = Dir(Range("dirRAtemplate").Value & "\*RAt.docx") ' ensure consistency with messages below
    Do While templateName$ <> ""
      If Left(templateName$, 1) <> "~" Then
        .ListRows.Add AlwaysInsert:=True
        nTemplates = nTemplates + 1
        .DataBodyRange(nTemplates, 1) = templateName$
      End If
      templateName$ = Dir
    Loop
End With
If nTemplates = 0 Then MsgBox ("Did not find any RA templates in " & Range("dirRAtemplate").Value & vbNewLine & "RA template names must end with RAt.docx; award templates must start with Awd and standard templates (autoloaded) must start with Std")
ExitHandler:
Application.ScreenUpdating = True
Exit Sub
ErrHandler:
MsgBox ("Error " & Err.Number & ":" & Err.Description & vbNewLine & "while trying to list templates.  Ensure template directory, " & Range("dirRAtemplate").Value & ", is accessible.")
Resume ExitHandler
End Sub

Sub Picker_dirRAtemplate()
Dim folderName As String
folderName = FolderPicker("Choose input folder containing RA templates *RAt.docx", Range("dirRAtemplate").Value)
If folderName <> "" Then Range("dirRAtemplate").Value = folderName
Call List_Templates
End Sub

Sub Picker_dirRAoutput()
Dim folderName As String
folderName = FolderPicker("Choose output folder for populated RA drafts (.docm)", Range("dirRAoutput").Value)
If folderName <> "" Then Range("dirRAoutput").Value = folderName
End Sub

Sub CheckRAFolders()
  If Len(Range("dirRAtemplate").Value) < 2 Then Range("dirRAtemplate").Value = ThisWorkbook.path
  If Len(Range("dirRAoutput").Value) < 2 Then Call Picker_dirRAoutput
End Sub
