VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Settings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Function Descriptor() As clsSheetDescriptor
' return descriptor for this sheet
    Set Descriptor = New clsSheetDescriptor
    With Descriptor
     .name = Me.name
     .description = "Settings: Rptserver password and EPSCoR states"
     .tip = "Update password on " & .name
     .tabtitle = "Workbook Settings"
     .note = "Settings for report server credentials (see your office admin) and EPSCoR states.  This sheet is protected to keep password masked."
     .helpText = "You can enter your division's report server user and, optionally, password here, and update when the password changes. (See your office admin) Can also update current EPSCoR states"
     .helpvideo = ""
     .phase = 0 ' never refreshed
     .order = 9999 ' last
     .uigroup = .order Mod 10
     .depends = ""
    End With
End Function

Private Sub Worksheet_Change(ByVal Target As Range)
' reset user_id and rptPassword if these cells change
    If Range(Target.Address).Row = Range("user_id").Row Then
        User_Id = ""
        rptPassword = ""
    End If
End Sub

