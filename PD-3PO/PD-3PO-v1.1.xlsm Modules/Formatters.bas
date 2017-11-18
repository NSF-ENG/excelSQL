Attribute VB_Name = "Formatters"
Sub ConditionalFormater()
' only for those with color.
  AllPIs.Range("c7").FormatConditions(1).Interior.Color = RGB(255, 244, 230)
  PropsOnPanels.Range("c7").FormatConditions(1).Interior.Color = RGB(255, 244, 230)
  SubAwd.Range("c7").FormatConditions(1).Interior.Color = RGB(255, 244, 230)
  Transfers.Range("c7").FormatConditions(2).Interior.Color = RGB(255, 244, 230)

End Sub

Sub AFormatter()
'
' Formatter Macro
    Selection.ClearContents
    Selection.Insert Shift:=xlDown
    Columns("A:A").Select
    Selection.Insert Shift:=xlToRight
    ActiveWindow.DisplayGridlines = False
    Range("B9").Select
    ActiveSheet.ListObjects(1).TableStyle = "TableFall"
    Call AHeader
End Sub

Sub AHeader()
Attribute AHeader.VB_ProcData.VB_Invoke_Func = " \n14"
'
' AHeader Macro
'
Dim ws As Worksheet
Set ws = ActiveSheet

    Rows("1:1").Select
    Selection.RowHeight = 40
    Rows("2:6").Select
    Selection.RowHeight = 16.5
    
    Sheets("Projects").Select
    Rows("1:4").Select
    Selection.Copy
    ws.Select
    Rows("1:1").Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    ws.Names.Add name:="run_datetime", RefersTo:="=" & ws.name & "!$C$2"
    ws.Names.Add name:="run_nrows", RefersTo:="=" & ws.name & "!$C$3"
    
End Sub

Sub ANamer()
'
' AHeader Macro
'
Dim ws As Worksheet
Set ws = ActiveSheet
    ws.Names.Add name:="run_datetime", RefersTo:="=" & ws.name & "!$C$2"
    ws.Names.Add name:="run_nrows", RefersTo:="=" & ws.name & "!$C$3"
    
End Sub


Sub Comments_Tom()
' format all comments on the active sheet ' from http://chandoo.org/wp/2009/09/11/format-comment-box/

 Dim MyComments As Comment
 Dim LArea As Long
 For Each MyComments In ActiveSheet.Comments
   With MyComments.Shape
   If .AutoShapeType <> msoShapeRoundedRectangle Then
    .AutoShapeType = msoShapeRoundedRectangle
    .TextFrame.Characters.Font.name = "Corbel"
    .TextFrame.Characters.Font.Bold = True
    .TextFrame.Characters.Font.Size = 10
    .TextFrame.Characters.Font.ColorIndex = 2
    .Line.ForeColor.RGB = RGB(0, 0, 0)
    .Line.BackColor.RGB = RGB(255, 255, 255)
    .Fill.Visible = msoTrue
    .Fill.ForeColor.RGB = RGB(27, 95, 109)
    .Fill.OneColorGradient msoGradientDiagonalUp, 1, 0.23
   End If
 End With
 Next 'comment
 End Sub
 
Sub CommentsToHelp()
 Dim c As Comment
 Call ClearTable(HelpTab.ListObjects("HelpPopup"))
 For Each c In InputTab.Comments
  With HelpTab.ListObjects("HelpPopup").ListRows.Add(AlwaysInsert:=True).Range
  .Cells(1, 1) = c.Parent.Address
  .Cells(1, 2) = c.Parent
  .Cells(1, 3) = c.Text
 End With
 Next c
End Sub



'Sub ld()
'Dim ws As Object
'Dim sd As clsSheetDescriptor
'
' Application.ScreenUpdating = False
' Application.EnableEvents = False
' Application.AutoCorrect.AutoFillFormulasInLists = False
'
' For Each ws In ThisWorkbook.Sheets
' On Error Resume Next ' do only the sheets with a Descriptor
'  Set sd = ws.Descriptor()
'  If Err = 0 Then 'have descriptor
'    On Error GoTo 0.Add(
'    With InputTab.ListObjects("MyTable").ListRows.Add(AlwaysInsert:=True).Range
'     .Cells(1, 1).Value = sd.name 'could use CodeName
'      .Cells(1, 2).Value = sd.description
'      .Cells(1, 3).Value = sd.usage
'    End With
'  ElseIf Err.Number <> 438 Then
'    If MsgBox("Error: " & Err.Number & " " & Err.description, vbOKCancel) <> vbOK Then End
'  End If
' Next ws
'
' On Error GoTo 0
'End Sub
'Public Function TabProperty(sheetName As String, Optional propertyName As String = "description") As String
'' Can be used to put registered description on a sheet; needs worksheet to have been saved to work.
'Dim cp As CustomProperty
'
'    TabProperty = "No " & propertyName
'    On Error GoTo errTabDesc
'    For Each cp In Worksheets(Mid(sheetName, InStrRev(sheetName, "]") + 1)).CustomProperties
'        If cp.Name = propertyName Then TabProperty = cp.Value
'    Next
'exitTabDesc:
'    On Error GoTo 0
'    Exit Function
'errTabDesc:
'    Resume exitTabDesc
'End Function


' for QueryStatusWatcher
'Dim SchedRecalc As Date
'Dim TicsLeft As Long

'Sub QueryStatusWatcher()
''recalculate sheet table status for after background query
' InputTab.Range("SheetTable[Status]").Calculate
' Call StartQueryStatusWatcher ' need to keep calling the timer, as the ontime only runs once
'End Sub
'
'Sub StartQueryStatusWatcher(Optional tics As Long = 0)
' TicsLeft = tics + TicsLeft - 1
' Debug.Print TicsLeft
' SchedRecalc = Now + TimeValue("00:00:05")
' If TicsLeft > 0 Then Application.OnTime SchedRecalc, "QueryStatusWatcher"
'End Sub
'
'Sub EndQueryStatusWatcher()
' On Error Resume Next
' Application.OnTime EarliestTime:=SchedRecalc, Procedure:="QueryStatusWatcher", Schedule:=False
' TicsLeft = 0
'End Sub
'
'
'Function QueryStatus(mySheet As Worksheet) As Variant
'' return number of rows in query table as status
'  QueryStatus = "..."
'  On Error Resume Next
'  With mySheet.ListObjects.Item(1)
'    If Not .QueryTable.Refreshing Then
'        If .DataBodyRange Is Nothing Then
'           QueryStatus = 0
'        Else
'           QueryStatus = .DataBodyRange.Rows.Count
'        End If
'    End If
'  End With
'  On Error GoTo 0
'End Function
'
'Function PivotStatus(mySheet As Worksheet) As Variant
'' return refresh time of last pivot table
'   PivotStatus = "..."
'   On Error Resume Next
'   With mySheet.PivotTables
'    PivotStatus = Format(.Item(.Count).RefreshDate, "yy-mm-dd hh:mm")
'   End With
'   On Error GoTo 0
'End Function


Private Sub shapeLister()
Dim ws As Worksheet
Dim shp, s2 As Shape
Dim t As Long

For Each ws In ThisWorkbook.Sheets
 'Debug.Print "Call " & ws.CodeName & ".Initialize"
 For Each shp In ws.Shapes
  t = 0
  On Error Resume Next
  t = shp.Type
  If t = 17 Then
    Debug.Print " Sheets(""" & ws.name;
    Debug.Print """).Shapes(""" & shp.name;
    Debug.Print """).Name = :: t:" & shp.Type;
    Debug.Print " h:" & shp.TextFrame2.HasText;
    Debug.Print " x:" & shp.TextFrame2.TextRange.Text;
    Debug.Print
   End If
    
    For Each s2 In shp.GroupItems
       t = 0
       t = s2.Type
       If t = 17 Then
        Debug.Print "     Sheets(""" & ws.name;
        Debug.Print """).Shapes(""" & s2.name;
        Debug.Print """).Name = :: t:" & s2.Type;
        Debug.Print " h:" & s2.TextFrame2.HasText;
        Debug.Print " x:" & s2.TextFrame2.TextRange.Text;
        Debug.Print
       End If
    Next s2
  Next shp
Next ws
On Error GoTo 0
End Sub


Sub PivotCacheReport()
'http://ramblings.mcpher.com/Home/excelquirks/snippets/pivotcache
Dim pc As PivotCache
Dim s As String, sn As String
Dim ws As Worksheet
Dim pt As pivotTable
With ActiveWorkbook
    For Each pc In .PivotCaches
        s = "Pivotcache " & CStr(pc.Index) & " uses " & CStr(pc.MemoryUsed) & " and has " _
        & CStr(pc.RecordCount) & " records"
        s = s & Chr(10) & "The following pivot tables use it"
        For Each ws In .Worksheets
            sn = ws.name
            For Each pt In ws.PivotTables
                If pt.CacheIndex = pc.Index Then
                    If Len(sn) > 0 Then
                        s = s & Chr(10) & sn & Chr(10)
                    sn = ""
                    End If
                    s = s & Replace(pt.name, "PivotTable", "PT") & ","
                End If
            Next pt
        Next ws
        Debug.Print (s)
    Next pc
    sn = Chr(10) & "Couldnt find the pivotcache for these pivot tables"
    s = ""
    For Each ws In .Worksheets
        For Each pt In ws.PivotTables
            If pt.CacheIndex < 1 Or pt.CacheIndex > .PivotCaches.Count Then
                s = s & Chr(10) & ws.name & ":" & Replace(pt.name, "PivotTable", "PT")
            End If
        Next pt
    Next ws
    If (Len(s) > 0) Then
        Debug.Print (sn & s)
    End If
End With
End Sub


Sub fixnames()
Dim nm As name
Dim i As Long
For Each nm In ThisWorkbook.Names
i = InStr(1, nm.name, "!run_")
If i > 0 Then
nm.name = Mid(nm.name, i + 1)
Debug.Print nm.name & ":" & nm.RefersTo
End If
Next nm
End Sub

'Sub ChangeNameScope()
'http://www.mrexcel.com/forum/excel-questions/665337-change-scope-defined-name-excel-2007-a.html#post3297886
'Dim nm As name, locNam
'Dim wSh As Worksheet
'Set wSh = ActiveSheet
'For Each nm In ActiveWorkbook.Names
'    If InStr(nm.RefersTo, wSh.name) > 0 Then
'        On Error Resume Next
'        If Not nm.RefersToRange Is Nothing Then
'            With nm.RefersToRange
'                locNam = nm.name
'                'remove global name
'                ActiveWorkbook.Names(nm.name).Delete
'                'Add local name
'                .Parent.Names.Add name:=locNam, RefersTo:="=" & .Address
'            End With
'        End If
'    On Error GoTo 0
'    End If
'Next nm
'End Sub

Sub PivotCacheClearRubbish()
Dim pc As PivotCache
Dim ws As Worksheet
With ActiveWorkbook
    For Each pc In .PivotCaches
        pc.MissingItemsLimit = xlMissingItemsNone
    Next pc
End With

For Each pc In ActiveWorkbook.PivotCaches
  On Error Resume Next
  pc.Refresh
Next pc
On Error GoTo 0
End Sub



'https://www.experts-exchange.com/articles/1457/Automate-Exporting-all-Components-in-an-Excel-Project.html
'Remember to add a reference to Microsoft Visual Basic for Applications Extensibility
'Exports all VBA project components containing code to a folder in the same directory as this spreadsheet.
Public Sub ExportAllComponents()
    Dim VBComp As VBIDE.VBComponent
    Dim destDir As String, fName As String, ext As String
    'Create the directory where code will be created.
    'Alternatively, you could change this so that the user is prompted
    If ActiveWorkbook.Path = "" Then
        MsgBox "You must first save this workbook somewhere so that it has a path.", , "Error"
        Exit Sub
    End If
    destDir = ActiveWorkbook.Path & "\" & ActiveWorkbook.name & " Modules"
    If Dir(destDir, vbDirectory) = vbNullString Then MkDir destDir
    
    'Export all non-blank components to the directory
    For Each VBComp In ActiveWorkbook.VBProject.VBComponents
        If VBComp.CodeModule.CountOfLines > 0 Then
            'Determine the standard extention of the exported file.
            'These can be anything, but for re-importing, should be the following:
            Select Case VBComp.Type
                Case vbext_ct_ClassModule: ext = ".cls"
                Case vbext_ct_Document: ext = ".cls"
                Case vbext_ct_StdModule: ext = ".bas"
                Case vbext_ct_MSForm: ext = ".frm"
                Case Else: ext = vbNullString
            End Select
            If ext <> vbNullString Then
                fName = destDir & "\" & VBComp.name & ext
                'Overwrite the existing file
                'Alternatively, you can prompt the user before killing the file.
                If Dir(fName, vbNormal) <> vbNullString Then Kill (fName)
                VBComp.Export (fName)
            End If
        End If
    Next VBComp
End Sub
