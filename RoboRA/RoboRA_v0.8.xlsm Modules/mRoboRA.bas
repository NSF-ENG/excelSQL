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
    s = VBA.Mid$(abstr, i, 5) & "|" & s
    Exit Do
  ElseIf Not VBA.Mid$(abstr, i - 1, 3) Like "[a-zA-Z][?][ '""]" Then
    i = i - 3
    s = VBA.Mid$(abstr, i, 8) & "|" & s
  End If
  i = InStrRev(abstr, "?", i - 1)
Loop
SummarizeQuestMarks = s
End Function

Private Sub test_SummarizeQuestMarks()
Debug.Print SummarizeQuestMarks("?Testing? Is this OK? ""And this?"" This houldn't be??  ?Done?.")
End Sub

Function autoPasteRA(IE As InternetExplorerMedium, prop_id As String, RA As String) As String
' stuff RA into text box using mAutocoder functions
Dim i As Integer, j As Integer
Dim overwriteQ As Variant

overwriteQ = Range("overwrite_option").Value
If (Len(prop_id) <> 7) Then ' warn that this is not a proposal id
    autoPasteRA = prop_id & " not a prop_id" & vbNewLine
    Exit Function
End If

IE.Navigate ("https://www.ejacket.nsf.gov/ej/showProposal.do?Continue=Y&ID=" & prop_id)
Call myWait(IE)
IE.Navigate ("https://www.ejacket.nsf.gov/ej/processReviewAnalysis.do?dispatch=add&uniqId=" & prop_id & VBA.LCase$(VBA.Left$(VBA.Environ$("USERNAME"), 7)))
Call myWait(IE)

If IE.Document.getElementsByName("text")(0) Is Nothing Then
  autoPasteRA = prop_id & " can't visit eJ RA" & vbNewLine
  Exit Function
End If

With IE.Document.getElementsByName("text")(0)
  .Focus
  If (Len(.Value) < 10) Or (overwriteQ = 3) Then
   .Focus
   .Value = RA
  ElseIf (overwriteQ = 2) Then ' ask permission to overwrite
    AppActivate Application.Caption
    DoEvents
    If (MsgBox("OK to overwrite existing RA for " & prop_id & vbNewLine & .Value, vbOKCancel) = vbOK) Then
     .Focus
     .Value = RA
    Else ' permission not granted
      autoPasteRA = prop_id & " not overwritten." & vbNewLine
      Exit Function
    End If
  Else ' never overwrite
    autoPasteRA = prop_id & " has text in RA field." & vbNewLine
    Exit Function
  End If
End With

Call myWait(IE)
If Not IE.Document.getElementsByName("save")(0) Is Nothing Then
  IE.Document.getElementsByName("save")(0).Click
  Call myWait(IE)
  autoPasteRA = ""
Else
  autoPasteRA = prop_id & " can't save eJ RA" & vbNewLine
End If
End Function

' Initialization and Preference states
' noSharedRAtemplate folder:  Fresh copy of RoboRA
'   - choose default, ListTemplates, & show splash
'\\collaboration.inside.nsf.gov@SSL\DavWWWRoot\eng\meritreview\SiteAssets\ENG Tools Websites and Best Practices\RoboRA\RAtemplates\*RAt.docx
Sub ckInitialization()
' This runs on workbook open
If Len(Range("dirSharedRAtemplate").Value) < 2 Then
  Prefs.Activate
#If Mac Then ' show mac instructions, and End.  Don't initialize
  Prefs.Range("WelcomeMac").Activate
  End
#Else ' initialize on PC
  Range("dirSharedRAtemplate").Value = "\\collaboration.inside.nsf.gov@SSL\DavWWWRoot\eng\meritreview\SiteAssets\ENG Tools Websites and Best Practices\RoboRA\RAtemplates\"
#End If
End If
Call List_Templates
End Sub


' List_Templates is called when we at least have the base (online) template folder name, even if it is not accessible at the moment.
' Prefer the local template folder name, if we have one, but if it contains no templates, offer to copy.
'
' Templates: personal if non-blank or base (should never be blank, but may be offline)
' Whenever personal templates change, listTemplates
'    if none, or  offer to renew from base.  (fail to renew, blank?  If no templates offer to renew.)
' refresh or renew personal templates?

Function folderRAtemplate() As String
' return the name of folder of local or base templates.
Dim dirRAtemplate As String
dirRAtemplate = fixEndSeparator(Range("dirRAtemplate").Value)
If Len(dirRAtemplate) < 2 Then dirRAtemplate = fixEndSeparator(Range("dirSharedRAtemplate").Value)
If Len(dirRAtemplate) < 2 Then
  Prefs.Activate
  MsgBox ("Please set RAtemplate folder on Prefs tab before continuing")
  End
End If
folderRAtemplate = dirRAtemplate
End Function

' do we neet to hangle path separators for http vs file?
Sub List_Templates() ' list RA templates available (used by data validation)
Dim templateName As String
Dim nTemplates As Integer
Dim dirRAtemplate As String
dirRAtemplate = folderRAtemplate()
nTemplates = 0
'Application.ScreenUpdating = False
On Error GoTo ErrHandler
With Prefs.ListObjects("AvailableTemplates")
  templateName$ = Dir(dirRAtemplate & "*RAt.docx") ' ensure consistency with messages below
  If templateName$ <> "" Then If Not .DataBodyRange Is Nothing Then .DataBodyRange.Delete
    Do While templateName$ <> ""
      If VBA.Left$(templateName$, 1) <> "~" Then
        .ListRows.Add AlwaysInsert:=True
        nTemplates = nTemplates + 1
        .DataBodyRange(nTemplates, 1) = templateName$
      End If
      templateName$ = Dir
    Loop
End With
Application.ScreenUpdating = True
If nTemplates = 0 Then
  If MsgBox("Did not find any RA templates in " & dirRAtemplate & "; shall I copy the standard templates to that folder?" _
           & vbNewLine & "Note: RA template names must end with RAt.docx; award templates must start with Awd and standard templates (autoloaded) must start with Std", vbOKCancel) = vbOK Then
    Call renewFiles("\\collaboration.inside.nsf.gov@SSL\DavWWWRoot\eng\meritreview\SiteAssets\ENG Tools Websites and Best Practices\RoboRA\RAtemplates\*.docx", dirRAtemplate)
    Call List_Templates
  End If
End If
ExitHandler:
Exit Sub
ErrHandler:
Application.ScreenUpdating = True
If Err.Number = 52 Then
MsgBox ("Cannot access template folder " & dirRAtemplate & vbNewLine & "I'll hope this is a network connection issue that will be fixed.")
Else
MsgBox ("Error " & Err.Number & ":" & Err.Description & vbNewLine & "while trying to list templates.  Ensure template folder, " & dirRAtemplate & ", is accessible.")
End If
Resume ExitHandler
End Sub

'Pickers for RA templates, RA output, and RoboRA location
'Note: RoboRA must be saved on a drive due to current limitations of MailMerge.


Sub Picker_dirSharedRAtemplate()
Dim folderName As String
folderName = FolderPicker("Choose folder containing base RA templates *RAt.docx", Range("dirSharedRAtemplate").Value)
If folderName <> "" Then Range("dirSharedRAtemplate").Value = folderName
Call List_Templates
End Sub

Sub Picker_dirRAtemplate()
Dim folderName As String
folderName = FolderPicker("Choose folder for personal RA templates *RAt.docx", Range("dirRAtemplate").Value)
If folderName <> "" Then Range("dirRAtemplate").Value = folderName
Call List_Templates
End Sub

Sub Picker_dirRAoutput()
Dim folderName As String
folderName = FolderPicker("Choose output folder for populated RA drafts", Range("dirRAoutput").Value)
If folderName <> "" Then Range("dirRAoutput").Value = folderName
End Sub






'Sub installMacros()
'Dim pathName As String
'
'On Error GoTo ErrHandler:
''JSS PC vs mac version
'pathName$ = "%appdata%\Microsoft\Word\STARTUP"
'If Dir(pathName$, vbDirectory) = "" Then MkDir (pathName$)
'MsgBox ("copying RAaddin.dotm into " & pathName$)
''JSS copy file RAaddin.dotm and trust it.
'ExitHandler:
'  On Error GoTo 0
'  Exit Sub
'ErrHandler:
'  MsgBox ("Error in Install_Raddin: " & Err.Number & ":" & Err.Description)
'  Resume ExitHandler
'End Sub



