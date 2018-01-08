VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Prefs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Worksheet_SelectionChange(ByVal Target As Range)
If Selection.count = 1 Then
    If Not (Intersect(Target, Prefs.Range("SharedRATemplateTable")) Is Nothing) Then
        Prefs.Range("dirSharedRAtemplate").Value = Prefs.Range("J" & Target.Row).Value
        List_Templates
    End If
End If
End Sub
