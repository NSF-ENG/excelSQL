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
Debug.Print SummarizeQuestMarks("?Testing? Is this OK? ""And this?"" This shouldn't be??  ?Done?.")
End Sub

Sub spRoboRApage()
  ' try to open the sharepoint roboRA page; ignore any errors, because we may not have access or it may be moved.
  On Error Resume Next
  ActiveWorkbook.FollowHyperlink Address:="https://collaboration.inside.nsf.gov/eng/meritreview/ENG%20Tools%20Websites%20and%20Best%20Practices/RoboRA.aspx"
  On Error GoTo 0
End Sub
Function strRAtemplateFolder() As String
'path to RAtemplate folder
With Prefs.Shapes("comboRAtemplateFolder").ControlFormat
  strRAtemplateFolder = ThisWorkbook.path & Application.PathSeparator & .List(.Value) & Application.PathSeparator
End With
End Function

Function IsRAtemplateFolderChosen() As Boolean
Dim fname As String
 IsRAtemplateFolderChosen = Prefs.Range("RAtemplateFolderIndex").Value > 0
 If IsRAtemplateFolderChosen Then
    fname = strRAtemplateFolder
    If Len(Dir(fname & "*RAt.docx")) < 2 Then
      Prefs.Activate
      MsgBox ("I did not find any RA templates in " & fname & vbNewLine & "Please ensure that there is an appropriate RAtemplates folder selected on Prefs #3 before continuing")
      End
    End If
 End If
End Function

Sub ckInitialization()
' This runs on workbook open
If Not IsRAtemplateFolderChosen() Then
    #If Mac Then ' show mac instructions, and End.  Don't initialize
      Prefs.Range("WelcomeMac").Activate
    #Else ' initialize on PC
      Prefs.Range("A1").Activate
    #End If
    End
End If
End Sub

Sub ckMailMerge(Optional needTemplates As Boolean = False)
' test all preconditions for a successful mail merge
Dim tmp As String
On Error GoTo errHandler
#If Mac Then
  MsgBox ("Mac users: Sorry, but this operation needs a reportserver connection and/or Word/IE automation libraries that are available only on a PC (including VDI/Citrix).")
  Prefs.Range("WelcomeMac").Activate
  End
#Else
  Dim path As String
  path = ThisWorkbook.path & Application.PathSeparator
  If VBA.Left$(path, 4) = "http" Then
    MsgBox ("RoboRA must be installed on a local (C:), personal (R:), or shared drive to perform mail merge and create review docs or RA drafts." & vbNewLine _
    & "Please install from the zip on sharepoint; Your version from " & path & " can be used for queries only.")
    spRoboRApage
    End
  End If
  If Len(Dir(path & "RoboRACleanCopy.dotm")) < 2 Or Len(Dir(path & "RAhelpTemplate.docx")) < 2 Then
    MsgBox ("Template files RoboRACleanCopy.dotm and RAhelpTemplate.docx must be in the folder with RoboRA" & vbNewLine _
    & "At least one is missing from " & path & ". Please install from the zip on sharepoint")
    spRoboRApage
    End
  End If
  If needTemplates Then
    If Not IsRAtemplateFolderChosen Then
      Prefs.Activate
      MsgBox ("Please first select your RAtemplate folder in Prefs #3 and return to RAdata to Make Indicated RAs.")
      End
    End If
  End If
#End If
ExitHandler:
  Exit Sub
errHandler:
  MsgBox ("Error " & Err.Number & " in ckMailMerge: " & Err.Description & vbNewLine _
  & "Attempting to continue...")
  Resume Next ' ExitHandler
End Sub

Sub Picker_RAoutput()
'pick folder for RA templates on active sheet
Dim folderName As String
folderName = FolderPicker("Choose output folder for populated RA drafts", Prefs.Range("RAoutput").Value)
If folderName <> "" Then
  Prefs.Range("RAoutput").Value = folderName
  RoboRA.Range("RAoutput").Value = folderName
  Advanced.Range("RAoutput").Value = folderName
End If
End Sub

Function goodFolderName(fname As String) As Boolean
'folder names to ignore in RoboRA
  goodFolderName = fname <> "" And VBA.Left(fname, 1) <> "."
End Function
  
Sub List_RAfolders()
' lists the folders in directory of thisworkbook
Dim path As String
Dim folderName As String
Dim nFolders As Integer
nFolders = 0
path = ThisWorkbook.path & Application.PathSeparator
If VBA.Left$(path, 4) = "http" Then
  MsgBox ("This copy of RoboRA is in " & path & ", but needs to be installed on a local (C:), personal (R:), or shared drive to populate templates. See Prefs #2, or install the zip from sharepoint.")
  Prefs.Activate
  spRoboRApage
  End
End If
Application.ScreenUpdating = False
On Error GoTo errHandler
With Prefs.ListObjects("FoldersWithRoboRA")
  folderName = Dir(path, vbDirectory)
  If Not .DataBodyRange Is Nothing Then .DataBodyRange.Delete
  Do While folderName <> ""
    If (VBA.Left$(folderName, 1) <> ".") And (vbDirectory = (VBA.GetAttr(path & folderName) And vbDirectory)) Then
      .ListRows.Add AlwaysInsert:=True
      nFolders = nFolders + 1
      .DataBodyRange(nFolders, 1) = VBA.Mid(folderName, VBA.InStrRev(Application.PathSeparator, folderName) + 1)
    End If
    folderName = Dir
  Loop
End With
Application.ScreenUpdating = True
If nFolders = 0 Then
  MsgBox ("No folders found in " & path & vbNewLine _
  & "Note: Any folders of RA template folders must be installed alongside RoboRA.xlsm in this folder. See Prefs #3.")
Else
  Prefs.Range("RAtemplateFolderIndex").Value = 1
  List_Templates
End If
ExitHandler:
Exit Sub
errHandler:
Application.ScreenUpdating = True
If Err.Number = 52 Then
  MsgBox ("Cannot access folder " & path & vbNewLine _
        & "I'll hope this is a network connection issue that will be fixed.")
Else
  MsgBox ("Error " & Err.Number & ":" & Err.Description & vbNewLine _
        & "while trying to list template folders. Ensure template folders are stored with RoboRA, currently in " & path & ".")
End If
Resume ExitHandler
End Sub

Sub List_Templates()
' list RA templates available (used by data validation)
If IsRAtemplateFolderChosen Then
    Dim templateName As String
    Dim nTemplates As Integer
    Dim dirRAtemplate As String
    dirRAtemplate = strRAtemplateFolder
    nTemplates = 0
    Application.ScreenUpdating = False
    On Error GoTo errHandler
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
      MsgBox ("Did not find any RA templates in " & strRAtemplateFolder & vbNewLine _
               & "Note: Template folder (selected in Prefs #3) must be with RoboRA in " & ThisWorkbook.path & vbNewLine _
               & "RA template names must end with RAt.docx; award templates must start with Awd and standard templates (autoloaded) must start with Std")
    End If
End If
ExitHandler:
Exit Sub
errHandler:
Application.ScreenUpdating = True
If Err.Number = 52 Then
  MsgBox ("Cannot access template folder " & strRAtemplateFolder & vbNewLine _
        & "I'll hope this is a network connection issue that will be fixed.")
Else
  MsgBox ("Error " & Err.Number & ":" & Err.Description & vbNewLine _
        & "while trying to list templates.  Ensure template folder, " & strRAtemplateFolder & ", is correct and accessible.")
End If
Resume ExitHandler
End Sub


'Sub Picker_dirRoboRA()
'Dim folderName As String
'folderName = FolderPicker("Choose folder on a drive (not SharePoint or OneDrive) to save RoboRA", "R:\")
'If folderName <> "" Then
''  If VBA.Left$(folderName, 4) = "http" Then
''    MsgBox ("Please choose a folder on a drive, and not with an http address (i.e. not SharePoint or OneDrive)")
''  Else
'    Prefs.Range("dirRoboRA").Value = folderName
'    ActiveWorkbook.SaveAs Filename:=fixEndSeparator(folderName) & ActiveWorkbook.name, _
'                        FileFormat:=xlOpenXMLWorkbookMacroEnabled, CreateBackup:=False
''  End If
'End If
'End Sub

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



