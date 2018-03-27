VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
#If Mac Then
Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)
Cancel = MsgBox("Are you sure you want to save RoboRA on a Mac?" & vbNewLine _
& "Only PCs (including VDI/Citrix) have native connection to ReportServer and the libraries needed to automate Word & IE." & vbNewLine _
& "If you save this for its data, you may need to get a fresh copy from Sharepoint before querying again.", vbOKCancel) <> vbOK
End Sub
#End If

Private Sub Workbook_Open()
  Application.Calculation = xlCalculationAutomatic
  Call ckInitialization
End Sub


