Attribute VB_Name = "mMailMerge"
Option Explicit

Sub MakeIndicatedRAs()
'derived from macro recording with assistance from several stackoverflow posts
'Uses RAtemplate column to decide action and RA template
'blank = ignore this one
'Awd = Award: refresh budget page?
'Std = Standard decline --  automatically stuff to eJacket

 Dim i As Integer
 Dim t As Integer
 Dim nRA As Integer
 Dim countRA As Integer
 Dim strWordDoc As Variant
 Dim path As String, strFilename As String
 Dim dirRAtemplate As String, strRAtemplate As String
 Dim dirRAoutput As String, strRAoutput As String
 Dim prop_id As String
 Dim warn As String
 Dim autoDeclineQ As Boolean, hasAuto As Boolean, newWordQ As Boolean
 Dim pt As PivotTable
 #If Mac Then
   MsgBox ("Making RA drafts needs to be done on a PC, including VDI/Citrix, in this version of RoboRA, because Macs don't have the libraries needed to Autoload to eJacket.")
   Prefs.Range("WelcomeMac").Activate
 #Else 'PC
 Dim wdApp As Object
 Dim wdDoc As Object
 Dim IE As InternetExplorerMedium
 warn = ""

Call ckMailMerge(True)
path = ThisWorkbook.path & Application.PathSeparator
dirRAtemplate = strRAtemplateFolder
dirRAoutput = fixEndSeparator(Prefs.Range("RAoutput").Value)
 
'check that templates exist for all actionable items.
'if any action is upload, check that eJ running.

For Each pt In HiddenSettings.PivotTables ' find templatesUsed pivot table and refresh
On Error Resume Next
If pt.name = "templatesUsed" Then Exit For
Next
If Not pt Is Nothing Then pt.RefreshTable
If pt Is Nothing Or Err.Number <> 0 Then
  MsgBox "Internal Error: Can't refresh pivot table templatesUsed on HiddenSettings tab."
  GoTo errHandler:
End If
On Error GoTo 0

nRA = 0
hasAuto = Prefs.CheckBoxes("cbAutoloadAll").Value = 1 ' do we have to open eJacket for autoload?
With RAData.Range("RADataQTable[RAtemplate]")
 For i = 1 To .Rows.count  ' quick check if there are any RAs to do
  strRAtemplate = Application.Trim(.Cells(i, 1))
  If Len(strRAtemplate) > 2 And strRAtemplate <> "(blank)" And (VBA.Left$(strRAtemplate, 2) <> "zz") Then
    nRA = nRA + 1 ' we have an RA to do
    If Not hasAuto Then hasAuto = (VBA.Left$(strRAtemplate, 3) = "Std") ' Look for first Std (Auto) decline
  End If
 Next i
End With
If nRA = 0 Then
    MsgBox ("On RAData, please select RAtemplates to indicate which RAs to prepare. If dropdowns in RAtemplate column are empty, select the RAtemplate folder in Prefs #3.")
    GoTo ExitHandler:
End If

Call renewFiles(path & "RoboRACleanCopy.dotm", dirRAoutput)
'If Prefs.CheckBoxes("cbConfirmActions").Value = 1 Then confirm ("About to start Mail Merge to create RA drafts")
ufProgress.Show vbModeless

If hasAuto Then Set IE = openEJacket()
    
On Error Resume Next ' connect to or start Word  'JSS mac version?
Set wdApp = GetObject(, "Word.Application")
newWordQ = wdApp Is Nothing ' Do we need to create a new wdApp?
If newWordQ Then Set wdApp = CreateObject("Word.Application")
On Error GoTo 0

' Sort by RecRkMin because our dummy line for formatting must come first.
   With RAData.ListObjects("RADataQTable").Sort
        .SortFields.Clear
        .SortFields.Add Key:=RAData.Range("RADataQTable[[#All],[RecRkMin]]"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortTextAsNumbers
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
countRA = 0
For t = 2 To pt.RowRange.count - 1 ' skip header and totals rows in pivot table
strRAtemplate = Application.Trim(pt.RowRange.Cells(t, 1))
 If Len(strRAtemplate) > 2 And strRAtemplate <> "(blank)" And (VBA.Left$(strRAtemplate, 2) <> "zz") Then ' we have an RA template
   Set wdDoc = wdApp.Documents.Open(dirRAtemplate & strRAtemplate)

    Do While wdDoc Is Nothing ' NOT TESTED
      If (MsgBox("can't find Word template " & dirRAtemplate & strRAtemplate & vbNewLine & " Open via dialog?", vbOKCancel) <> vbOK) Then GoTo ExitHandler:
      Dim fd As FileDialog 'File Picker dialog box.
      Set fd = Application.FileDialog(msoFileDialogFilePicker)
        With fd
           .AllowMultiSelect = False
           If .Show <> -1 Then GoTo ExitHandler: 'Show File Picker; abort on cancel
           strWordDoc = .SelectedItems(1)
           Set wdDoc = wdApp.Documents.Open(Filename:=strWordDoc, ReadOnly:=True)
       End With
       Set fd = Nothing
    Loop
    On Error GoTo 0
    
    autoDeclineQ = (Prefs.CheckBoxes("cbAutoloadAll").Value = 1) Or (VBA.Left$(strRAtemplate, 3) = "Std")
    wdDoc.Activate
    wdApp.Visible = True
    DoEvents
    wdDoc.ActiveWindow.View.ReadingLayout = False ' avoid windows '13 reading layout default
        
    With RAData.Range("RADataQTable[RAtemplate]") ' need RAfname as next column!!!
      For i = 2 To .Rows.count ' do the RAs, skipping the first
        If strRAtemplate = Application.Trim(.Cells(i, 1)) Then ' we have an RA to do
        countRA = countRA + 1
        UpdateProgressBar (countRA / (nRA + 1))
        strRAoutput = dirRAoutput & Application.Trim(.Cells(i, 2)) & VBA.Format$(Now, "yymmdd_hhmm") & ".docx" ' make output file name
    '    Application.ScreenUpdating = False
    '    Application.DisplayAlerts = False
       With wdDoc.MailMerge
           .MainDocumentType = wdFormLetters
          
          .OpenDataSource name:=ThisWorkbook.FullName, _
              LinkToSource:=False, AddToRecentFiles:=False, Revert:=False, Format:=wdOpenFormatAuto, _
              Connection:="Data Source='" & ThisWorkbook.FullName & "';Mode=Read", _
              SQLStatement:="SELECT * FROM `RAData$`"
     
          .Destination = wdSendToNewDocument
          .SuppressBlankLines = True
            
           With .DataSource
             .FirstRecord = i
             .LastRecord = i
           End With 'data source
          .Execute Pause:=True 'False
        End With 'mail merge
        With wdApp.ActiveDocument
          If autoDeclineQ Then
            Dim RAtext As String
            With .ActiveWindow.Selection
              .WholeStory
              RAtext = FixIPSText(StripDoubleBrackets(.Text))
              .Collapse
            End With ' selection
            prop_id = Application.Trim(RAData.Range("RADataQTable[[prop_id0]]").Cells(i, 1).Value)
            
            warn = warn & autoPasteRA(IE, prop_id, RAtext)
            .ReadOnlyRecommended = True
          End If
          .AttachedTemplate = dirRAoutput & "RoboRACleanCopy.dotm" 'All the macros are in this template
          .SaveAs2 Filename:=strRAoutput, FileFormat:=wdFormatXMLDocument, LockComments:=False, Password:="", AddToRecentFiles _
            :=True, WritePassword:="", ReadOnlyRecommended:=False, EmbedTrueTypeFonts _
            :=False, SaveNativePictureFormat:=False, SaveFormsData:=False, _
            SaveAsAOCELetter:=False
          '.SaveAs Filename:=strRAoutput, FileFormat:=wdFormatXMLDocumentMacroEnabled, _
           '        AddToRecentFiles:=True, ReadOnlyRecommended:=False
          .Close savechanges:=wdSaveChanges
          End With 'document
       ' ActiveWindow.Close
      End If ' done mailmerge
     Next i
     End With 'table range
  End If
   If Not (wdDoc Is Nothing) Then
     wdDoc.Close savechanges:=wdDoNotSaveChanges
     Set wdDoc = Nothing
   End If
 Next t

ExitHandler:
Unload ufProgress
If hasAuto Then Call closeEJacket(IE)
If Not (wdDoc Is Nothing) Then
   wdDoc.Close savechanges:=wdDoNotSaveChanges
   Set wdDoc = Nothing
End If
If newWordQ And Not (wdApp Is Nothing) Then
  wdApp.Quit
  Set wdApp = Nothing
End If
If warn <> "" Then
  activateApp
  MsgBox ("Warnings copied to clipboard: " & vbNewLine & warn)
  CopyText (warn)
End If
Exit Sub

errHandler:
  MsgBox ("Error in MakeIndicatedRAs: " & Err.Number & ":" & Err.Description)
  Resume ExitHandler
#End If 'PC
End Sub

Sub makeProjText()
'derived from macro recording with assistance from several stackoverflow posts

Dim path As String, strWordDoc As String, strPDFOutputName As String
Dim dirRAtemplate As String, dirRAoutput As String
Dim newWordQ As Boolean
' #If Mac Then
'   MsgBox ("Functions that do mail merge need to be run on a PC, including VDI/Citrix.")
'   Prefs.Range("WelcomeMac").Activate
' #Else 'PC
 Dim wdApp As Object
 Dim wdDoc As Object
 
ckMailMerge
path = ThisWorkbook.path & Application.PathSeparator
dirRAtemplate = strRAtemplateFolder
dirRAoutput = fixEndSeparator(Prefs.Range("RAoutput").Value)
strWordDoc = path & "RAhelpTemplate.docx"
strPDFOutputName = dirRAoutput & "RAhelp" & VBA.Format$(Now(), "_yymmdd_hhmm")
 
ufProgress.Show vbModeless

On Error Resume Next
Set wdApp = GetObject(, "Word.Application")
newWordQ = wdApp Is Nothing ' Do we need to create a new wdApp?
If newWordQ Then Set wdApp = CreateObject("Word.Application")
On Error GoTo 0
 
'    Application.ScreenUpdating = False
'    Application.DisplayAlerts = False
Call UpdateProgressBar(0.05)

 Set wdDoc = wdApp.Documents.Open(Filename:=strWordDoc, ReadOnly:=True)
 wdDoc.Activate
 wdApp.Visible = True
 DoEvents
 wdDoc.ActiveWindow.View.ReadingLayout = False

Call UpdateProgressBar(0.1)
'Connection:= "Provider=Microsoft.ACE.OLEDB.12.0;User ID=Admin;Data Source=C:\Users\Jack Snoeyink\Desktop\tmp.xlsm';Mode=Read;Extended Properties=""HDR=YES;IMEX=1;"";Jet OLEDB:System database="""";Jet OLEDB:Registry Path="""";Jet OLEDB:Engine Type=3"
    With wdDoc.MailMerge
       .MainDocumentType = 0 'wdFormLetters, wdOpenFormatAuto
       .OpenDataSource name:=ThisWorkbook.FullName, _
          LinkToSource:=False, AddToRecentFiles:=False, Revert:=False, Format:=0, _
          Connection:="Data Source='" & ThisWorkbook.FullName & "';Mode=Read" _
          , SQLStatement:="SELECT * FROM `ProjText$`"
 
        .Destination = 0 'wdSendToNewDocument
        .SuppressBlankLines = True
        With .DataSource
            .FirstRecord = 1
            .LastRecord = -16
        End With
        .Execute Pause:=True 'False
    End With
Call UpdateProgressBar(0.6)
    'export format pdf=17, opt for screen=1,wdExportCreateHeadingBookmarks=1
    wdApp.ActiveDocument.ExportAsFixedFormat OutputFileName:=strPDFOutputName, ExportFormat:= _
        17, OpenAfterExport:=True, OptimizeFor:= _
        1, Range:=0, from:=1, To:=1, _
        Item:=0, IncludeDocProps:=True, KeepIRM:=True, _
        CreateBookmarks:=1, DocStructureTags:=True, _
        BitmapMissingFonts:=True, UseISO19005_1:=False
Call UpdateProgressBar(0.9)
ExitHandler:
 wdApp.ActiveDocument.Close savechanges:=0 ' don't save changes
 wdDoc.Close savechanges:=0
 Set wdDoc = Nothing
 If newWordQ Then Set wdApp = Nothing
 Unload ufProgress
 Exit Sub
errHandler:
'if err.Number = 4605 then msgbox("Word complains about opening in reading; please see Prefs #7.")
End Sub
