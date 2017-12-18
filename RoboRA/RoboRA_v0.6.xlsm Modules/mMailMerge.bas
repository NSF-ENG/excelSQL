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
 Dim wdApp As Object
 Dim wdDoc As Object
 Dim strWordDoc As Variant
 Dim strThisWorkbook As String, strOutputPath As String, strFilename As String, strRAtemplate As String, strRAoutput As String
 Dim dirRAtemplate As String, dirRAoutput As String
 Dim prop_id As String
 Dim warn As String
 Dim autoDeclineQ As Boolean, hasAuto As Boolean
 Dim IE As InternetExplorerMedium
 Dim pt As PivotTable
warn = ""

strThisWorkbook = ThisWorkbook.FullName
dirRAtemplate = Range("dirRAtemplate").Value
If VBA.Right$(dirRAtemplate, 1) <> Application.pathSeparator Then dirRAtemplate = dirRAtemplate & Application.pathSeparator
dirRAoutput = Range("dirRAoutput").Value
If VBA.Right$(dirRAoutput, 1) <> Application.pathSeparator Then dirRAoutput = dirRAoutput & Application.pathSeparator

'If Not checkRoboRAFolders Then Exit Sub
'check that templates exist for all actionable items.
'if any action is upload, check that eJ running.

For Each pt In HiddenSettings.PivotTables ' find templatesUsed pivot table and refresh
On Error Resume Next
If pt.name = "templatesUsed" Then Exit For
Next
If Not pt Is Nothing Then pt.RefreshTable
If pt Is Nothing Or Err.Number <> 0 Then
  MsgBox "Can't refresh pivot table templatesUsed on HiddenSettings tab."
  GoTo ErrHandler:
End If
On Error GoTo 0

nRA = 0
hasAuto = RoboRA.CheckBoxes("cbAutoloadAll").Value = 1
With Range("RADataTable[RAtemplate]")
 For i = 1 To .Rows.count  ' quick check
  strRAtemplate = Application.Trim(.Cells(i, 1))
  If Len(strRAtemplate) > 2 And strRAtemplate <> "(blank)" And (VBA.Left$(strRAtemplate, 2) <> "zz") Then
    nRA = nRA + 1 ' we have an RA to do
    If Not hasAuto Then hasAuto = (VBA.Left$(strRAtemplate, 3) = "Std") ' Look for first Std (Auto) decline
  End If
 Next i
End With
If nRA = 0 Then
    MsgBox ("On RAData, please select RAtemplates to indicate which RAs to prepare. If dropdown in RAtemplate column is empty, pick the RAtemplate folder on the Advanced tab.")
    GoTo ExitHandler:
End If

Call renewFiles("\\collaboration.inside.nsf.gov@SSL\DavWWWRoot\eng\meritreview\SiteAssets\ENG Tools Websites and Best Practices\RoboRA\RoboRACleanCopy.dotm", dirRAoutput)
'If RoboRA.CheckBoxes("cbConfirmActions").Value = 1 Then confirm ("About to start Mail Merge to create RA drafts")
ufProgress.Show vbModeless

If hasAuto Then Set IE = openEJacket()
    
On Error Resume Next ' start Word  'JSS mac version?
Set wdApp = GetObject(, "Word.Application")
If wdApp Is Nothing Then
    Set wdApp = CreateObject("Word.Application")
End If
On Error GoTo 0

' Sort by RecRkMin because our dummy line for formatting must come first.
   With RAData.ListObjects("RADataTable").Sort
        .SortFields.Clear
        .SortFields.Add Key:=Range("RADataTable[[#All],[RecRkMin]]"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortTextAsNumbers
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
           Set wdDoc = wdApp.Documents.Open(strWordDoc)
       End With
       Set fd = Nothing
    Loop
    On Error GoTo 0
    
    autoDeclineQ = (RoboRA.CheckBoxes("cbAutoloadAll").Value = 1) Or (VBA.Left$(strRAtemplate, 3) = "Std")
    wdDoc.Activate
    wdApp.Visible = True
    
        
    With Range("RADataTable[RAtemplate]") ' need RAfname as next column!!!
      For i = 2 To .Rows.count ' do the RAs, skipping the first
        If strRAtemplate = Application.Trim(.Cells(i, 1)) Then ' we have an RA to do
        countRA = countRA + 1
        UpdateProgressBar (countRA / (nRA + 1))
        strRAoutput = dirRAoutput & Application.Trim(.Cells(i, 2)) & VBA.Format$(Now, "yymmdd_hhmm") & ".docm" ' make output file name
    '    Application.ScreenUpdating = False
    '    Application.DisplayAlerts = False
       With wdDoc.MailMerge
           .MainDocumentType = wdFormLetters
          
          .OpenDataSource name:=strThisWorkbook, _
              LinkToSource:=False, AddToRecentFiles:=False, Revert:=False, Format:=wdOpenFormatAuto, _
              Connection:="Data Source='" & strThisWorkbook & "';Mode=Read", _
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
            prop_id = Application.Trim(Range("RADataTable[[prop_id0]]").Cells(i, 1).Value)
            
            warn = warn & autoPasteRA(IE, prop_id, RAtext)
            .ReadOnlyRecommended = True
          End If
          .AttachedTemplate = dirRAoutput & "RoboRACleanCopy.dotm" 'JSS what if this is on a different computer?
          .SaveAs2 Filename:=strRAoutput, FileFormat:=wdFormatXMLDocumentMacroEnabled, LockComments:=False, Password:="", AddToRecentFiles _
            :=True, WritePassword:="", ReadOnlyRecommended:=False, EmbedTrueTypeFonts _
            :=False, SaveNativePictureFormat:=False, SaveFormsData:=False, _
            SaveAsAOCELetter:=False
          '.SaveAs Filename:=strRAoutput, FileFormat:=wdFormatXMLDocumentMacroEnabled, _
           '        AddToRecentFiles:=True, ReadOnlyRecommended:=False
          .Close SaveChanges:=wdSaveChanges
          End With 'document
       ' ActiveWindow.Close
      End If ' done mailmerge
     Next i
     End With 'table range
  End If
   If Not (wdDoc Is Nothing) Then
     wdDoc.Close SaveChanges:=wdDoNotSaveChanges
     Set wdDoc = Nothing
   End If
 Next t

ExitHandler:
Unload ufProgress
If hasAuto Then Call closeEJacket(IE)
If Not (wdDoc Is Nothing) Then
   wdDoc.Close SaveChanges:=wdDoNotSaveChanges
   Set wdDoc = Nothing
End If
If warn <> "" Then
  AppActivate Application.Caption
  DoEvents
  MsgBox ("Warnings copied to clipboard: " & vbNewLine & warn)
  CopyText (warn)
End If
Exit Sub

ErrHandler:
  MsgBox ("Error in MakeIndicatedRAs: " & Err.Number & ":" & Err.Description)
  Resume ExitHandler
End Sub

Sub makeProjText()
'derived from macro recording with assistance from several stackoverflow posts

 Dim wdApp As Object, wdDoc As Object
 Dim strWordDoc As String, strThisWorkbook As String, strPDFOutputName As String
 Dim dirRAtemplate As String, dirRAoutput As String
 
 dirRAtemplate = Advanced.Range("dirRAtemplate").Value
If VBA.Right$(dirRAtemplate, 1) <> Application.pathSeparator Then dirRAtemplate = dirRAtemplate & Application.pathSeparator
dirRAoutput = Advanced.Range("dirRAoutput").Value
If VBA.Right$(dirRAoutput, 1) <> Application.pathSeparator Then dirRAoutput = dirRAoutput & Application.pathSeparator
 
 strThisWorkbook = ThisWorkbook.FullName
 strWordDoc = dirRAtemplate & "RAhelpTemplate.docx"
 strPDFOutputName = dirRAoutput & "RAhelp" & VBA.Format$(Now(), "_yymmdd_hhmm")
 
ufProgress.Show vbModeless

On Error Resume Next
Set wdApp = GetObject(, "Word.Application")
If wdApp Is Nothing Then
    Set wdApp = CreateObject("Word.Application")
End If
On Error GoTo 0
 
'    Application.ScreenUpdating = False
'    Application.DisplayAlerts = False
Call UpdateProgressBar(0.05)

 Set wdDoc = wdApp.Documents.Open(strWordDoc)
 wdDoc.Activate
 wdApp.Visible = True

Call UpdateProgressBar(0.1)
'Connection:= "Provider=Microsoft.ACE.OLEDB.12.0;User ID=Admin;Data Source=C:\Users\Jack Snoeyink\Desktop\tmp.xlsm';Mode=Read;Extended Properties=""HDR=YES;IMEX=1;"";Jet OLEDB:System database="""";Jet OLEDB:Registry Path="""";Jet OLEDB:Engine Type=3"
    With wdDoc.MailMerge
       .MainDocumentType = 0 'wdFormLetters, wdOpenFormatAuto
       .OpenDataSource name:=strThisWorkbook, _
          LinkToSource:=False, AddToRecentFiles:=False, Revert:=False, Format:=0, _
          Connection:="Data Source='" & strThisWorkbook & "';Mode=Read" _
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
 wdApp.ActiveDocument.Close SaveChanges:=0 ' don't save changes
 wdDoc.Close SaveChanges:=0
 Set wdDoc = Nothing
 Unload ufProgress
End Sub
