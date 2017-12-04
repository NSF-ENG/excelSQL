Attribute VB_Name = "mUtilities"
Option Explicit
' general utilities PD-3PO family of tools, Jack Snoeyink, Oct 2017

Sub CleanUpSheet(ws As Worksheet, Optional emptyRow As Long = 7)
' delete blank rows on sheet below lowest listObject range
' pass emptyrow as the index of the first row that could be empty.
' Avoid blanks past a table for mail merge.
' Also keeps sheets from growing too tall (saving memory and preserving stability).
Dim i As Long, r As Long, lastRow As Long
With ws
    On Error Resume Next
    For i = 1 To ws.ListObjects.count
      r = ws.ListObjects(i).Range.End(xlDown).Row + 1
      If emptyRow < r Then emptyRow = r
    Next i
    
    For i = 1 To ws.PivotTables.count
      r = ws.PivotTables(i).RowRange.End(xlDown).Row + 1
      If emptyRow < r Then emptyRow = r
    Next i
    On Error GoTo 0
    lastRow = ws.UsedRange.Rows.count
    If emptyRow < lastRow Then ws.Rows(emptyRow & ":" & lastRow).Delete
End With
End Sub

Sub RefreshPivotTables(ws As Worksheet, qt As QueryTable)
' refreshing pivot tables that are tied to a given query table  (PD-3PO only?)
 Dim PT As PivotTable
 For Each PT In ws.PivotTables
   PT.PivotTableWizard SourceType:=xlDatabase, SourceData:=qt.ListObject.Name
   If Not (PT Is Nothing) Then PT.RefreshTable
  Next
End Sub

Sub ClearTable(LO As ListObject)
  With LO
    If Not .DataBodyRange Is Nothing Then .DataBodyRange.Delete
  End With
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

Function pathSeparator() As String
'path separators differ on Mac and PC
#If Mac Then
    pathSeparator = "/"
#Else
    pathSeparator = "\"
#End If
End Function

Function wordAddinPath() As String
' This is where VBA Add-ins go on Mac and PC
#If Mac Then
    wordAddinPath = "/Library/Application Support/Microsoft/Office365/User Content/Startup/Word"
#Else
    wordAddinPath = Environ("AppData") & "\Microsoft\Word\STARTUP"
#End If
End Function


Function FolderPicker(title As String, Optional initFolder As String = vbNullString) As String
'
With Application.FileDialog(msoFileDialogFolderPicker)
    .title = title
    .InitialFileName = initFolder & pathSeparator()
    .Show
    If .SelectedItems.count > 0 Then FolderPicker = .SelectedItems(1)
End With
End Function
Function confirm(msg As String, Optional abortQ As Boolean = False) As Integer
' Allow user to confirm action.  vbCancel, or vbNo with abort=True will call End, aborting calling Sub.
' Otherwise you can check the return value =vbNo to skip the action
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
    Dim separator As String, s As String
    separator = pathSeparator() ' Mac or PC
    arrPath = Split(path, separator)
    s = arrPath(LBound(arrPath)) & separator
    For i = LBound(arrPath) + 1 To UBound(arrPath)
        s = s & arrPath(i) & separator
        If Dir(s, vbDirectory) = "" Then
          MkDir s
        End If
    Next
End Sub
Sub renewFiles(from As String, topath As String, Optional verbosity As Integer = 1)
' copy files matching from (include filter *.* or *RAt.docx or *.do*) that have been updated or that don't exist in topath
' old files get "backup"datetime added, so nothing is lost.
Dim FSO As Object
Dim fromdate As Date, todate As Date
Dim frompath As String, fName As String, separator As String
separator = pathSeparator() ' mac or PC
Set FSO = CreateObject("scripting.filesystemobject")
If Not FSO.FolderExists(topath) Then
  Call confirm("Create folder " & topath, True) ' abort if vbCancel or vbNo
  createPath (topath)
End If
If Right$(topath, 1) <> separator Then topath = topath & separator ' add separator if needed
frompath = Left(from, InStrRev(from, separator))
fName = Dir(from) ' get first matching file
While fName <> ""
    fromdate = FileDateTime(frompath & fName)
    On Error Resume Next
    todate = FileDateTime(topath & fName)
    Select Case Err.Number
        Case 0 ' file exists
            If fromdate <= todate Then GoTo nextFile ' File does not need to be renewed
            If verbosity > 0 Then
               If confirm("Update file " & topath & fName) = vbNo Then GoTo nextFile
            End If
            On Error GoTo 0
            FSO.moveFile Source:=topath & fName, Destination:=topath & fName & "backup" & Format(Now, "yyyy-mm-dd hh-mm-ss")
            FSO.copyFile Source:=frompath & fName, Destination:=topath & fName
        Case 53: ' file not found; create it
            If verbosity > 1 Then
              If confirm("Install file " & topath & fName) = vbNo Then GoTo nextFile
            End If
            On Error GoTo 0
            FSO.copyFile Source:=frompath & fName, Destination:=topath & fName
        Case 76: ' path not found, even though we must have tried to create it.  Abort
            MsgBox ("To path " & topath & " could not be created; aborting.")
            End
        Case Else ' other error
            MsgBox ("Error:" & Err.Number & ":" & Err.Description & " in renewFiles. Skipping: " & topath & fName)
            GoTo nextFile
    End Select

nextFile:
  fName = Dir ' get next matching file
Wend

Set FSO = Nothing
End Sub

Sub test_renewFiles()
Call renewFiles("R:\RATemplates\*.docx", "R:\Temp\Sub\Level2")
End Sub

