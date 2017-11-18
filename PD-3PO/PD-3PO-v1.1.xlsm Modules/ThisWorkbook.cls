VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

Option Explicit

Private Sub Workbook_Open()
' Initialize query tables for the sheets that have them.
  Call InitializeAllQueryTables
  Call hideSheets
' Always open on the Input tab
    Worksheets("Input").Activate
End Sub

Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)
' unhide before saving to avoid pulling variables from hidden sheets on open.
  Call UnhideAllSheets
End Sub

Private Sub Workbook_AfterSave(ByVal Success As Boolean)
' hide again after saving
  Call hideSheets
End Sub

