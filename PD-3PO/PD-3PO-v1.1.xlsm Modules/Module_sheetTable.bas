Attribute VB_Name = "Module_sheetTable"
Option Explicit
' Utility functions for working with sheet descriptors and the SheetTable interface
' Every refreshable sheet needs to have RefreshRS and Descriptor subs.

' This module provides methods to work with and update an interface SheetTable
' that selects and manages query and pivot table sheets.
' It stores no data itself; that all lives in the table.
'
' On the sheet holding the table, you want to support hiding with this _Change event handler
' Private Sub Worksheet_Change(ByVal Target As Range)
'  If Not Intersect(Target, Me.Range("SheetTable[Ref?]")) Is Nothing Then hideSheets
' End Sub

' SheetTable fields: 1:Status(datetime), 2:Num, 3:Limit, 4:Ref (n/H), 5:Tab (Name), 6:Description, 7:phase, (8:order)
' Phase 1 & 2 are for queries that others (esp pivot tables) depend on.
' Pivot tables must be in phases after the tables they depend on.
' Long queries should be in phase 3 (default).
Enum stColumnEnum
  emai = 1
  sname
  ref
  limit
  num
  stat
  descr
  phase
  order
End Enum

Sub Refresh()
'refresh All Tabs that depend on InputProps
Dim i, phase As Long

If Not QTInitializedFlag Then Call InitializeAllQueryTables

Call pwdHandler.doPwd

Call preQuery(InputTab)
If (InputTab.Range("CountTable[[num_retrieved]]").Value > 300) Then _
   If MsgBox("Warning: This will retrieve at least " & InputTab.Range("CountTable[[num_retrieved]]").Value _
             & " proposals per tab.", vbOKCancel) <> vbOK Then End

With InputTab.Range("SheetTable") ' assumed to exist
   For phase = 1 To 3 ' do queries in 3 phases, syncing in between
   OutstandingQueries = 3 - phase ' start at zero only for phase 3
   .Calculate
    For i = 1 To .Rows.Count
     If phase = Val(.Cells(i, stColumnEnum.phase).Value) And Len(Trim(.Cells(i, stColumnEnum.ref).Value)) < 1 Then ' correct phase & want to refresh
       If IsNumeric(.Cells(i, stColumnEnum.limit).Value) And .Cells(i, stColumnEnum.limit).Value > 0 Then ' query with TOP limit
        Call ThisWorkbook.Sheets(.Cells(i, stColumnEnum.sname).Value).RefreshRS(" TOP " & .Cells(i, stColumnEnum.limit).Value & " ")
       Else
        Call ThisWorkbook.Sheets(.Cells(i, stColumnEnum.sname).Value).RefreshRS
       End If
     End If
    Next i
   If phase < 3 Then Application.CalculateUntilAsyncQueriesDone
  Next phase
  .Calculate
End With
End Sub

Sub InitializeAllQueryTables()
' Run by workbookopen; need to run after error or adding sheets
' Processes worksheets as objects because not all sheets have Sub InitializeQueryTable
Dim ws As Object
    For Each ws In ThisWorkbook.Worksheets
      On Error Resume Next
      Call ws.InitializeQueryTable ' initializes for those sheets that have this function
    Next ws
On Error GoTo 0
QTInitializedFlag = True
End Sub

Sub ClearSheets()
' Run by workbookopen; need to run after error or adding sheets
' Processes worksheets as objects because not all sheets have Sub InitializeQueryTable
If MsgBox("This wlll clear data from both visible and hidden sheets.", vbOKCancel) <> vbOK Then End

If Not QTInitializedFlag Then Call InitializeAllQueryTables

Dim ws As Object
    For Each ws In ThisWorkbook.Worksheets
      On Error Resume Next
      Call ws.ClearRS ' clears those sheets that have this function
    Next ws
On Error GoTo 0
'Call PivotCacheClearRubbish
End Sub

Sub hideSheets()
' hide sheets if first letter is H
' Clear and convert to H if first letter is C
' Delete tab and interface row if first letter is D

Dim i As Long
Dim s, mySheet As String
Dim eflag, sflag, didDelete As Boolean

sflag = Application.ScreenUpdating
If sflag Then Application.ScreenUpdating = False
eflag = Application.EnableEvents
If eflag Then Application.EnableEvents = False
 'Debug.Print "H" & eflag & sflag ' JSS
   With InputTab.Range("SheetTable")
     If .Rows.Count > 1 Then
        For i = 1 To .Rows.Count
          s = UCase$(Left$(Trim$(.Cells(i, stColumnEnum.ref).Value), 1))
          mySheet = .Cells(i, stColumnEnum.sname).Value
          While s = "X" ' delete tab and row
            If Len(mySheet) > 1 Then ' have a tab
              didDelete = False
              On Error Resume Next
              didDelete = ThisWorkbook.Sheets(mySheet).Delete ' true for successful deletion
              On Error GoTo 0
              If didDelete Then
                .Rows(i).Delete ' delete interface row
                If i > .Rows.Count Then GoTo hideSheetExit
                s = UCase$(Left$(Trim$(.Cells(i, stColumnEnum.ref).Value), 1))
                mySheet = .Cells(i, stColumnEnum.sname).Value
              Else
                s = "C" ' delete failed or cancelled; hide instead
              End If
            End If
          Wend
          If Len(mySheet) > 1 Then ' have a tab
             If s = "C" Then ' clear the results then hide the tab
               s = "H"
               .Cells(i, stColumnEnum.ref).Value = "H"
               On Error Resume Next
               Call ThisWorkbook.Sheets(mySheet).ClearRS
               On Error GoTo 0
             End If
             On Error Resume Next ' Just in case someone renames sheets
             ThisWorkbook.Sheets(mySheet).Visible = (s <> "H")
         On Error GoTo 0
          End If
        Next i
     End If
    End With
hideSheetExit:
    InputTab.Visible = xlSheetVisible
    InputTab.Activate
    If eflag Then Application.EnableEvents = True
    If sflag Then Application.ScreenUpdating = True
    Exit Sub
End Sub

Sub UnhideAllSheets()
' unhide all sheets; for programming convenience
Dim ws As Worksheet

    Application.ScreenUpdating = False
     For Each ws In ThisWorkbook.Sheets
          ws.Visible = True
        Next ws
    InputTab.Activate
    Application.ScreenUpdating = True
End Sub

Private Sub InitializeAll()
' this is just for convenience while coding;
Dim ws As Worksheet
Dim shp, s2 As Shape
Dim t As Long

For Each ws In ThisWorkbook.Sheets
 'Debug.Print "Call " & ws.CodeName & ".Initialize"
 For Each shp In ws.Shapes
  t = 0
  On Error Resume Next
  t = shp.Type
  If t = 6 Then
    With shp.GroupItems
     Debug.Print ws.name & ",""" & .Shapes("tabTitle").TextFrame2.TextRange.Text _
                      & """, """ & .Shapes("tabNote").TextFrame2.TextRange.Text & """"
    End With
   End If
  Next shp
  With ws
     Debug.Print ws.name & ",""" & .Shapes("tabTitle").TextFrame2.TextRange.Text _
                      & """, """ & .Shapes("tabNote").TextFrame2.TextRange.Text & """"
    End With
Next ws
On Error GoTo 0
'Call Totals.initialize
'Call Orphans.initialize
'Call NewInst.initialize
'Call Projects.initialize
'Call AllPIs.initialize
'Call Panels.initialize
'Call Panelists.initialize
'Call SugRevr.initialize
'Call Reviewers.initialize
'Call SubAwd.initialize
'Call Transfers.initialize
'Call PropsOnPanels.initialize
'Call DDOverview.initialize
'Call DDReviews.initialize
'Call PropTracking.initialize
End Sub

Private Sub LabelTab(ws As Worksheet, title As String, note As String)
' look for shapes in group or isolated and label them
Dim shp, s2 As Shape
Dim t As Long

 For Each shp In ws.Shapes
  t = 0
  On Error Resume Next
  t = shp.Type
  If t = 6 Then ' group type
    With shp.GroupItems
        .Shapes("tabTitle").TextFrame2.TextRange.Text = title
        .Shapes("tabNote").TextFrame2.TextRange.Text = note
    End With
   End If
  Next shp
  With ws
    .Shapes("tabTitle").TextFrame2.TextRange.Text = title
    .Shapes("tabNote").TextFrame2.TextRange.Text = note
  End With
On Error GoTo 0
End Sub


Sub makeSheetTable()
' this clears the SheetTable and rebuilds it
Dim ws As Object
Dim sd As clsSheetDescriptor
Dim col, tint As Variant
Dim addr As String

If MsgBox("Do you really want to rebuild the SheetTable?", vbOKCancel) <> vbOK Then Exit Sub
 
 Application.ScreenUpdating = False
 Application.EnableEvents = False
 Application.AutoCorrect.AutoFillFormulasInLists = False

'tab color palette
col = Array(4, 4, 3, 12, 8, 5, 6, 6, 10, 9)
tint = Array(0, -0.25, -0.25, -0.5, -0.5, -0.25, -0.5, -0.35, -0.25, -0.25)

Call ClearTable(InputTab.ListObjects("SheetTable")) ' clear sheet table
Call ClearTable(HelpTab.ListObjects("HelpTable")) ' clear help table

 For Each ws In ThisWorkbook.Sheets
 On Error Resume Next ' do only the sheets with a Descriptor
  Set sd = ws.Descriptor()
  If Err = 0 Then 'have descriptor
  'Debug.Print ws.name
    On Error GoTo 0
    Call LabelTab(ws, UCase$(sd.tabtitle), sd.note)
    
    With InputTab.ListObjects("SheetTable").ListRows.Add(AlwaysInsert:=True).Range
      '.Cells(1, stColumnEnum.stat).Formula = "=" & sd.name & "!run_datetime"
      addr = ""
      On Error Resume Next
      addr = sd.name & "!" & Sheets(sd.name).Range("run_datetime").Address
      On Error GoTo 0
     .Cells(1, stColumnEnum.stat).Formula = ""
      If Len(addr) > 0 Then .Cells(1, stColumnEnum.stat).Formula = "=IFERROR(IF(NOW()-datevalue(" & addr & ")<1,RIGHT(" & addr & ",11),LEFT(" & addr & ",8)), " & addr & ")"
      On Error Resume Next
      .Cells(1, stColumnEnum.num).Formula = ""
      .Cells(1, stColumnEnum.num).Formula = "=" & sd.name & "!" & Sheets(sd.name).Range("run_nrows").Address
      On Error GoTo 0
      If sd.phase = 3 Then .Cells(1, stColumnEnum.ref) = "H"
     '.Cells(1, 5).Value = sd.name 'could use CodeName
       Call InputTab.Hyperlinks.Add(.Cells(1, stColumnEnum.sname), "", sd.name & "!$A$1", sd.tip, sd.name)
    With .Cells(1, stColumnEnum.sname).Font 'Color sheet Table
        .name = "Calibri"
        .FontStyle = "Bold"
        .Size = 12
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = 0
        .ThemeFont = xlThemeFontMinor
    End With
    With .Cells(1, stColumnEnum.sname).Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = col(sd.uigroup)
        .TintAndShade = tint(sd.uigroup)
        .PatternTintAndShade = 0
    End With
      .Cells(1, stColumnEnum.descr).Value = sd.description
      'Call InputTab.Hyperlinks.Add(.Cells(1, stColumnEnum.descr), "", sd.name & "!$A$1", sd.tip, sd.description)
   
      .Cells(1, stColumnEnum.phase).Value = sd.phase
      .Cells(1, stColumnEnum.order).Value = sd.order ' can delete/hide order column after table is built
    End With
    
    
    
    With ActiveWorkbook.Sheets(sd.name).Tab ' color Tab
        .ThemeColor = col(sd.uigroup)
        .TintAndShade = tint(sd.uigroup)
    End With
    
    With HelpTab.ListObjects("HelpTable").ListRows.Add(AlwaysInsert:=True).Range
      .Cells(1, 1) = sd.order
      '.Cells(1, 2) = sd.name
      Call HelpTab.Hyperlinks.Add(.Cells(1, 2), "", sd.name & "!$A$1", sd.tip, sd.name)
        With .Cells(1, 2).Font 'Color sheet Table
            .name = "Calibri"
            .FontStyle = "Bold"
            .Size = 12
            .Strikethrough = False
            .Superscript = False
            .Subscript = False
            .OutlineFont = False
            .Shadow = False
            .Underline = xlUnderlineStyleNone
            .ThemeColor = xlThemeColorDark1
            .TintAndShade = 0
            .ThemeFont = xlThemeFontMinor
        End With
        With .Cells(1, 2).Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .ThemeColor = col(sd.uigroup)
            .TintAndShade = tint(sd.uigroup)
            .PatternTintAndShade = 0
        End With
      .Cells(1, 3) = sd.tabtitle
      .Cells(1, 4) = sd.helpText
    End With
    
  ElseIf Err.Number <> 438 Then
    If MsgBox("Error: " & Err.Number & " " & Err.description, vbOKCancel) <> vbOK Then End
  End If
 Next ws
 
 On Error GoTo 0
  With InputTab.ListObjects("SheetTable").Sort  ' sort by order hints; can delete order column after table is built
      .SortFields.clear
      .SortFields.Add Key:=InputTab.Range("SheetTable[[#Headers],[Order]]"), SortOn:=xlSortOnValues, _
                  order:=xlAscending, DataOption:=xlSortTextAsNumbers
      .Header = xlYes
      .MatchCase = False
      .Orientation = xlTopToBottom
      .SortMethod = xlPinYin
      .Apply
  End With
  
  With HelpTab.ListObjects("HelpTable").Sort  ' sort by order hints; can delete order column after table is built
      .SortFields.clear
      .SortFields.Add Key:=HelpTab.Range("HelpTable[[#Headers],[Order]]"), SortOn:=xlSortOnValues, _
                  order:=xlAscending, DataOption:=xlSortTextAsNumbers
      .Header = xlYes
      .MatchCase = False
      .Orientation = xlTopToBottom
      .SortMethod = xlPinYin
      .Apply
  End With

  Call InitializeAllQueryTables
  Call hideSheets
  Application.EnableEvents = True
  Application.ScreenUpdating = True
  Application.AutoCorrect.AutoFillFormulasInLists = True
End Sub


Sub Mail_Sheets()
'modified from http://www.rondebruin.nl/win/winmail
'Working in Excel 2000-2016
'For Tips see: http://www.rondebruin.nl/win/winmail/Outlook/tips.htm
    Dim FileExtStr As String
    Dim FileFormatNum As Long
    Dim Sourcewb As Workbook
    Dim Destwb As Workbook
    Dim TempFilePath As String
    Dim TempFileName As String
    Dim OutApp As Object
    Dim OutMail As Object
    Dim sh As Worksheet
    Dim TheActiveWindow As Window
    Dim TempWindow As Window
    Dim i, j, nrow As Long
    Dim sheetName() As String

    With Application
        .ScreenUpdating = False
        .EnableEvents = False
    End With
    
    Call UnhideAllSheets
    
    With InputTab.Range("sheetTable")
      nrow = .Rows.Count
      ReDim sheetName(nrow)
      j = 0
      sheetName(j) = InputTab.name
      For i = 1 To nrow
        If UCase$(Left$(.Cells(i, stColumnEnum.emai), 1)) = "Y" Then
          j = j + 1
          sheetName(j) = .Cells(i, stColumnEnum.sname)
        End If
      Next i
    End With
    If j = 0 Then Exit Sub
    ReDim Preserve sheetName(0 To j)

    Set Sourcewb = ActiveWorkbook
    'Copy the sheets to a new workbook
    'We add a temporary Window to avoid the Copy problem if there is a List or Table in one of the sheets and if the sheets are grouped
    With Sourcewb
        Set TheActiveWindow = ActiveWindow
        Set TempWindow = .NewWindow
        TempFilePath = Environ$("temp") & "\"
        TempFileName = "Part of " & .name & " " & Format(Now, "dd-mmm-yy h-mm-ss")
'        .Theme.ThemeColorScheme.Save (TempFilePath & "PD-3POThemeColors.xml")
     ' we need the theme colors, which are on a protected workbook, so we hard code them here
    '
'    Debug.Print "Array(";
'         For i = 1 To .Theme.ThemeColorScheme.Count
'          Debug.Print "," & .Theme.ThemeColorScheme.Colors(i);
'        Next i
'        Debug.Print ")"
         Dim themeColors
         themeColors = Array(0, 16777215, 5526612, 12566463, 13810240, 47610, 2341776, 553198, 10466074, 4012501, 553198, 2499495)
        .Sheets(sheetName).Copy
    End With

    'Close temporary Window
    TempWindow.Close
    Set Destwb = ActiveWorkbook
    FileExtStr = ".xlsx"     'Set file extension/format to lose the macros
    FileFormatNum = 51
    '    'Change all cells in the worksheets to values if you want
    '    For Each sh In Destwb.Worksheets
    '        sh.Select
    '        With sh.UsedRange
    '            .Cells.Copy
    '            .Cells.PasteSpecial xlPasteValues
    '            .Cells(1).Select
    '        End With
    '        Application.CutCopyMode = False
    '    Next sh

    'Save the new workbook/Mail it/Delete it
    
    Set OutApp = CreateObject("Outlook.Application")
    Set OutMail = OutApp.CreateItem(0)

    With Destwb

        '.Theme.ThemeColorScheme.Load (TempFilePath & "PD-3POThemeColors.xml")
        For i = LBound(themeColors) To UBound(themeColors)
           .Theme.ThemeColorScheme.Colors(i + 1) = themeColors(i)
        Next i
        
        Dim xConnect As Object ' delete connections
        On Error Resume Next
        For Each xConnect In .Connections
        If xConnect.name <> "ThisWorkbookDataModel" Then xConnect.Delete
        Next xConnect
        On Error GoTo 0
        
        .Sheets(1).Visible = False ' hide input tab
        .Application.DisplayAlerts = False ' we are stripping the macros
        .SaveAs TempFilePath & TempFileName & FileExtStr, FileFormat:=FileFormatNum
        .Application.DisplayAlerts = True
        On Error Resume Next
        With OutMail
            .To = ""
            .CC = ""
            .BCC = ""
            .Subject = "Workbook (data only) " & Format(Now(), "short date")
            .Body = "Attached please find the worksheets from PD-3PO; hidden input tab has the parameters."
            .Attachments.Add Destwb.FullName
            .Display
        End With
        On Error GoTo 0
        .Close savechanges:=False
    End With

    'Delete the file you have send
    Kill TempFilePath & TempFileName & FileExtStr

    Set OutMail = Nothing
    Set OutApp = Nothing
    Call hideSheets
    With Application
        .ScreenUpdating = True
        .EnableEvents = True
    End With
End Sub

