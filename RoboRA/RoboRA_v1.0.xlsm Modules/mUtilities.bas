Attribute VB_Name = "mUtilities"
Option Explicit
' general utilities PD-3PO family of tools, Jack Snoeyink, Oct 2017

Sub CleanUpSheet(ws As Worksheet, Optional emptyRow As Long = 4)
' delete blank rows on sheet below lowest listObject range
' pass emptyrow as the index of the first row that could be empty.
' Avoid blanks past a table for mail merge.
' Also keeps sheets from growing too tall (saving memory and preserving stability).
Dim tb As ListObject
Dim pt As PivotTable
Dim r As Long, lastRow As Long

On Error Resume Next
For Each tb In ws.ListObjects
  If Not tb.TotalsRowRange Is Nothing Then
    r = tb.TotalsRowRange.Row + 1
  ElseIf tb.DataBodyRange Is Nothing Then
    r = tb.InsertRowRange.Row
  Else
    r = tb.ListRows(tb.ListRows.count).Range.Row + 1
  End If
  If emptyRow < r Then emptyRow = r
Next

For Each pt In ws.PivotTables
  If Not pt.RowRange Is Nothing Then
    r = pt.RowRange.Rows(pt.RowRange.count).Row + 1
    If emptyRow < r Then emptyRow = r
  End If
Next

emptyRow = emptyRow + 1
lastRow = ws.UsedRange.Rows.count
Application.DisplayAlerts = False
If emptyRow < lastRow Then ws.Rows(emptyRow & ":" & lastRow).Delete
Application.DisplayAlerts = True
On Error GoTo 0
End Sub
Sub RefreshPivotTables(ws As Worksheet)
' refreshing pivot tables on worksheet
 Dim pt As PivotTable
 For Each pt In ws.PivotTables
   If Not (pt Is Nothing) Then pt.RefreshTable
  Next
End Sub

Sub ClearTable(lo As ListObject)
  With lo
    If Not .DataBodyRange Is Nothing Then
      Application.DisplayAlerts = False
      .DataBodyRange.Delete
      Application.DisplayAlerts = True
    End If
  End With
End Sub

Sub ClearMatchingTables(t As String, Optional ws As Worksheet = Nothing)
' use wildcards to match table names to clear items in activesheet.
' note: probably want to use something that doesn't depend on activesheet.
If ws Is Nothing Then ws = ActiveSheet
Dim lo As ListObject
For Each lo In ws.ListObjects
  If (lo.name Like t) Then Call ClearTable(lo)
Next
End Sub
Sub PivotCacheClearRubbish()
' not used because it will remove timeline slicers
Dim pc As PivotCache
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
Sub ClearQTables()
Dim ws As Worksheet
Dim lo As ListObject
For Each ws In ThisWorkbook.Sheets
  Call ClearMatchingTables("*QTable", ws)
  Call RefreshPivotTables(ws)
  ' be careful not to clean up sheets with info below tables.
  If ws.name <> "HiddenSettings" And ws.name <> "Advanced" And ws.name <> "Prefs" Then Call CleanUpSheet(ws)
Next
Call PivotCacheClearRubbish
End Sub

'lFunction wordAddinPath() As String
' This is where VBA Add-ins go on Mac and PC
' not used
'#If Mac Then
'    wordAddinPath = "/Library/Application Support/Microsoft/Office365/User Content/Startup/Word"
'#Else
'    wordAddinPath = VBA.Environ$("AppData") & "\Microsoft\Word\STARTUP"
'#End If
'End Function

' Try to handle path separators on Mac & PC

Function fixEndSeparator(s As String) As String
'fix or add separator to end of name
Select Case VBA.Right$(s, 1)
 Case "/", "\": fixEndSeparator = VBA.Left$(s, Len(s) - 1) & Application.PathSeparator
 Case Else: fixEndSeparator = s & Application.PathSeparator
End Select
End Function

Function fixSeparators(s As String) As String
' convert mac or pc separators to current
fixSeparators = Replace(Replace(s, "/", Application.PathSeparator), "\", Application.PathSeparator)
End Function

Function fixSeparatorsAddEnd(s As String) As String
' convert mac or pc separators and add one at end if needed.
Select Case VBA.Right$(s, 1)
 Case "/", "\": s = VBA.Left$(s, Len(s) - 1) & Application.PathSeparator
 Case Else: s = s & Application.PathSeparator
End Select
fixSeparatorsAddEnd = Replace(Replace(s, "/", Application.PathSeparator), "\", Application.PathSeparator)
End Function

Sub test_fixEndSeparator()
  Debug.Print fixEndSeparator("r:\back\")
  Debug.Print fixEndSeparator("r:/fwd/")
  Debug.Print fixEndSeparator("r:\no")
  Debug.Print fixEndSeparator("r:")
  Debug.Print fixEndSeparator("")
End Sub

Sub test_fixSeparatorsAddEnd()
  Debug.Print fixSeparatorsAddEnd("r:\back\")
  Debug.Print fixSeparatorsAddEnd("r:/fwd/")
  Debug.Print fixSeparatorsAddEnd("r:\no")
  Debug.Print fixSeparatorsAddEnd("r:/")
  Debug.Print fixSeparatorsAddEnd("r:\")
  Debug.Print fixSeparatorsAddEnd("")
End Sub

Function FolderPicker(title As String, Optional initFolder As String = vbNullString) As String
'Mac version derived thanks to https://www.rondebruin.nl/mac/mac017.htm
#If Mac Then
    Dim RootFolder As String
    Dim scriptstr As String
    On Error Resume Next
    RootFolder = MacScript("return (path to desktop folder) as String")
    If Val(Application.Version) < 15 Then
        scriptstr = "(choose folder with prompt """ & title & """" _
           & " default location alias """ & RootFolder & """) as string"
    Else
        scriptstr = "return posix path of (choose folder with prompt """ & title & """" _
            & " default location alias """ & RootFolder & """) as string"
    End If
    FolderPicker = MacScript(scriptstr)
    On Error GoTo 0
#Else 'PC
With Application.FileDialog(msoFileDialogFolderPicker)
    .title = title
    .InitialFileName = fixEndSeparator(initFolder)
    .Show
    If .SelectedItems.count > 0 Then FolderPicker = .SelectedItems(1)
End With
#End If 'PC
End Function

Function confirm(msg As String, Optional abortQ As Boolean = False) As Integer
' Allow user to confirm action.  vbCancel, or vbNo with abort=True will call End, aborting calling Sub.
' Otherwise you can check the return value =vbNo to skip the action
activateApp
confirm = MsgBox(msg, vbYesNoCancel)
If confirm <> vbYes And (confirm = vbCancel Or abortQ) Then
    MsgBox ("Aborting action: please recheck parameters before initiating action.")
    End
End If
End Function

Sub createPath(path As String)
' make all directories on path if needed
' needs error handling
    Dim i As Long
    Dim arrPath As Variant
    Dim s As String
    arrPath = Split(fixSeparators(path), Application.PathSeparator)
    s = arrPath(LBound(arrPath)) & Application.PathSeparator
    For i = LBound(arrPath) + 1 To UBound(arrPath)
        s = s & arrPath(i) & Application.PathSeparator
        If Dir(s, vbDirectory) = "" Then
          MkDir s
        End If
    Next
End Sub

' Mac porting: this one may be problematic, and isn't used often. Refactor?
' this was to get files from sharepoint, but they tell me that webDAV will not be supported,
' so its only use is to refresh the CleanCopy macros in an RAoutput folder. Feb 2018
Sub renewFiles(from As String, topath As String, Optional verbosity As Integer = 1)
' copy files matching from (include filter *.* or *RAt.docx or *.do*) that have been updated or that don't exist in topath
' old files get "backup"datetime added, so nothing is lost.
' unused
Dim FSO As Object
Dim fromdate As Date, todate As Date
Dim frompath As String, fname As String
Set FSO = CreateObject("scripting.filesystemobject")
If Not FSO.FolderExists(topath) Then
  Call confirm("Create folder " & topath, True) ' abort if vbCancel or vbNo
  createPath (topath)
End If
topath = fixEndSeparator(topath)
frompath = VBA.Left$(from, VBA.InStrRev(from, Application.PathSeparator))
fname = Dir(from) ' get first matching file
While fname <> ""
    fromdate = FileDateTime(frompath & fname)
    On Error Resume Next
    todate = FileDateTime(topath & fname)
    Select Case Err.Number
        Case 0 ' file exists
            If fromdate <= todate Then GoTo nextFile ' File does not need to be renewed
            If verbosity > 0 Then
               If confirm("Update file " & topath & fname) = vbNo Then GoTo nextFile
            End If
            On Error GoTo 0
            FSO.moveFile Source:=topath & fname, Destination:=topath & fname & "backup" & VBA.Format$(Now, "yymmdd_hhmm")
            FSO.copyFile Source:=frompath & fname, Destination:=topath & fname
        Case 53: ' file not found; create it
            If verbosity > 1 Then
              If confirm("Install file " & topath & fname) = vbNo Then GoTo nextFile
            End If
            On Error GoTo 0
            FSO.copyFile Source:=frompath & fname, Destination:=topath & fname
        Case 76: ' path not found, even though we must have tried to create it.  Abort
            MsgBox ("To path " & topath & " could not be created; aborting.")
            End
        Case Else ' other error
            MsgBox ("Error:" & Err.Number & ":" & Err.Description & " in renewFiles. Skipping: " & topath & fname)
            GoTo nextFile
    End Select

nextFile:
  fname = Dir ' get next matching file
Wend

Set FSO = Nothing
End Sub

Sub test_renewFiles()
Call renewFiles("R:\RATemplates\*.docx", "R:\Temp\Sub\Level2")
End Sub

