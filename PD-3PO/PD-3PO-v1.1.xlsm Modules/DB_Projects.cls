VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DB_Projects"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit
Private QT As clsQueryTable

Function Descriptor() As clsSheetDescriptor
' return descriptor for this sheet
    Set Descriptor = New clsSheetDescriptor
    With Descriptor
     .name = Me.name
     .description = "Dashboard for Projects"
     .tip = "Unhide to Jump to " & .name
     .tabtitle = .description
     .note = "Pivot tables for Project data; refreshed automatically."
     .helpText = "Pivot tables show projects received, status, etc, allowing filters on Dir, Div, PO, request, status, and object class code."
     .helpvideo = ""
     .phase = 2
     .order = 105
     .uigroup = .order Mod 10
     .depends = Projects.name
    End With
End Function

'Sub ClearRS()
'  'Call RefreshPivotTables(Me)
'End Sub

Sub RefreshRS(Optional limit As String = "")
' refresh all pivot tables.  Must be in a later phase than tables this depends on.
  Me.Range("run_datetime").Value = "..."
  Call RefreshPivotTables(Me)
End Sub
