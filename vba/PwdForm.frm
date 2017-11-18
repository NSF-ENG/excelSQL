VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} PwdForm 
   Caption         =   "CredentialsForm"
   ClientHeight    =   3060
   ClientLeft      =   104
   ClientTop       =   432
   ClientWidth     =   4592
   OleObjectBlob   =   "PwdForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "PwdForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
' Jack Snoeyink      Oct 2, 2017
'assumes clsPwd has been instantiated in pwdHandler
Private Sub cmdPwdCancel_Click()
  Unload Me
  End ' abort calling program
End Sub

'Save userid & password in connections, and on ReportServerPwd tab if checkbox = true
Private Sub cmdPwdOK_Click()
    If Me.CheckBox1.Value = True Then ' save in ReportServerPwd or not
         ReportServerPwd.txtUserId.Value = Me.txtUserId.Value
         ReportServerPwd.txtPwd.Value = Me.txtPassword.Value
    Else
         ReportServerPwd.txtPwd.Value = ""
    End If
    Call FixConnections(Me.txtUserId.Value, Me.txtPassword.Value)
    'PwdForm.Hide
    Unload Me
End Sub

