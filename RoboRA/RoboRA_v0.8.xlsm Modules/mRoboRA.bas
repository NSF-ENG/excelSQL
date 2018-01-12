Attribute VB_Name = "mRoboRA"
Option Explicit
' Utility routines specific to RoboRA
Public Function SummarizeQuestMarks(abstr As String) As String
' Summarize ? from string that may have been converted from quotes, dashes, or other special characters.
Dim i As Long
Dim s As String
s = " "
i = VBA.InStrRev(abstr, "?")
Do While i > 0
  If i < 4 Then
    s = VBA.Mid$(abstr, i, 5) & "|" & s
    Exit Do
  ElseIf Not VBA.Mid$(abstr, i - 1, 3) Like "[a-zA-Z][?][ '""]" Then
    i = i - 3
    s = VBA.Mid$(abstr, i, 8) & "|" & s
  End If
  i = VBA.InStrRev(abstr, "?", i - 1)
Loop
SummarizeQuestMarks = s
End Function

Private Sub test_SummarizeQuestMarks()
Debug.Print SummarizeQuestMarks("?Testing? Is this OK? ""And this?"" This houldn't be??  ?Done?.")
End Sub


' Initialization and Preference states
' noSharedRAtemplate folder:  Fresh copy of RoboRA
'   - choose default, ListTemplates, & show splash
'\\collaboration.inside.nsf.gov@SSL\DavWWWRoot\eng\meritreview\SiteAssets\ENG Tools Websites and Best Practices\RoboRA\RAtemplates\*RAt.docx
Sub ckInitialization()
' This runs on workbook open
If Len(Prefs.Range("dirSharedRAtemplate").Value) < 2 Then
  Prefs.Activate
#If Mac Then ' show mac instructions, and End.  Don't initialize
  Prefs.Range("WelcomeMac").Activate
  End
#Else ' initialize on PC
  Prefs.Range("dirSharedRAtemplate").Value = "\\collaboration.inside.nsf.gov@SSL\DavWWWRoot\eng\meritreview\SiteAssets\ENG Tools Websites and Best Practices\RoboRA\RAtemplates\"
  Call List_Templates
#End If
End If
End Sub

Sub ckRAFolders()
Dim tmp As String
If VBA.Left$(ActiveWorkbook.FullName, 4) = "http" Then
  MsgBox ("RoboRA must be saved on a drive before attempting Mail Merge.  (See Prefs tab #3)")
  End
End If
tmp = folderRAtemplate()
If Len(Dir(tmp & "*RAt.docx")) < 2 Then
   Prefs.Activate
   MsgBox ("I did not find any RA templates in " & tmp & vbNewLine & "Please ensure that there is an appropriate RAtemplates folder selected on Prefs #2 before continuing")
   End
End If
tmp = Prefs.Range("dirRAoutput").Value
If Len(tmp) < 2 Then
  MsgBox ("Please select a folder for the output pdf & RA drafts")
  Call Picker_dirRAoutput
  If Len(Prefs.Range("dirRAoutput").Value) < 2 Then End
End If
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
  MsgBox ("Please select RAtemplates folder on Prefs tab (#2) before continuing")
  End
End If
folderRAtemplate = dirRAtemplate
End Function

' do we need to change path separators for http vs file?
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
  MsgBox ("Cannot access template folder " & dirRAtemplate & vbNewLine _
        & "I'll hope this is a network connection issue that will be fixed.")
Else
  MsgBox ("Error " & Err.Number & ":" & Err.Description & vbNewLine _
        & "while trying to list templates.  Ensure template folder, " & dirRAtemplate & ", is accessible.")
End If
Resume ExitHandler
End Sub

'Pickers for RA templates, RA output, and RoboRA location
'Note: RoboRA must be saved on a drive due to current limitations of MailMerge.

Sub Picker_dirRoboRA()
Dim folderName As String
folderName = FolderPicker("Choose folder on a drive (not SharePoint or OneDrive) to save RoboRA", "R:\")
If folderName <> "" Then
  If VBA.Left$(folderName, 4) = "http" Then
    MsgBox ("Please choose a folder on a drive, and not with an http address (i.e. not SharePoint or OneDrive)")
  Else
    Prefs.Range("dirRoboRA").Value = folderName
    ActiveWorkbook.SaveAs Filename:=fixEndSeparator(folderName) & ActiveWorkbook.Name, _
                        FileFormat:=xlOpenXMLWorkbookMacroEnabled, CreateBackup:=False
  End If
End If
End Sub

Sub Picker_dirSharedRAtemplate()
Dim folderName As String
folderName = FolderPicker("Choose folder containing base RA templates *RAt.docx", Prefs.Range("dirSharedRAtemplate").Value)
If folderName <> "" Then
  Prefs.Range("dirSharedRAtemplate").Value = folderName
  Call List_Templates
End If
End Sub

Sub Picker_dirRAtemplate()
Dim folderName As String
folderName = FolderPicker("Choose folder for personal RA templates *RAt.docx", Prefs.Range("dirRAtemplate").Value)
If folderName <> "" Then
  Prefs.Range("dirRAtemplate").Value = folderName
  Call List_Templates
End If
End Sub

Sub Picker_dirRAoutput()
Dim folderName As String
folderName = FolderPicker("Choose output folder for populated RA drafts", Prefs.Range("dirRAoutput").Value)
If folderName <> "" Then
  Prefs.Range("dirRAoutput").Value = folderName
  RoboRA.Range("dirRAoutput2").Value = folderName
  Advanced.Range("dirRAoutput3").Value = folderName
End If
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



