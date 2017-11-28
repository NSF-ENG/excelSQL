VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True


Private Sub Workbook_Open()
  If Len(Range("dirRAtemplate").Value) < 2 Then Range("dirRAtemplate").Value = ThisWorkbook.path
  If Len(Range("dirRAoutput").Value) < 2 Then Range("dirRAoutput").Value = Range("dirRAtemplate").Value
End Sub
