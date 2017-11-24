VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} PwdForm 
   Caption         =   "CredentialsForm"
   ClientHeight    =   3064
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
' Jack Snoeyink      Nov 24, 2017
Private Sub cmdPwdCancel_Click()
  Unload Me
  End ' abort calling program
End Sub

'Save userid & password in connections, and on ReportServerPwd tab if checkbox = true
Private Sub cmdPwdOK_Click()
    If Me.CheckBox1.Value = True Then ' save in ReportServerPwd or not
         HiddenSettings.Range("user_id").Value = Me.txtUserId.Value
         HiddenSettings.Range("rpt_pwd").Value = Me.txtPassword.Value
    Else
         HiddenSettings.Range("rpt_pwd").Value = ""
    End If
    Call FixConnections(Me.txtUserId.Value, Me.txtPassword.Value)
    'PwdForm.Hide
    Unload Me
End Sub

Private Sub UserForm_Click()

End Sub
