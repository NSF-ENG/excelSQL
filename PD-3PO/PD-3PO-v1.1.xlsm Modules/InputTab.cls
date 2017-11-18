VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "InputTab"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub Worksheet_Change(ByVal Target As Range)
'Debug.Print "C"; ' JSS
  On Error Resume Next
  If Not Intersect(Target, Range("SheetTable[Ref?]")) Is Nothing Then hideSheets
End Sub

Private Sub Retrieve_Click()
    Call Refresh
End Sub

Private Sub ClearResults_Click()
   Call ClearSheets
End Sub


