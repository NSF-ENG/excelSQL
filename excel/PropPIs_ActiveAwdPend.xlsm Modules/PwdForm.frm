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
  HiddenSettings.Range("rpt_pwd").Value = ""
  Me.Hide
  'Unload Me
  End ' abort calling program
End Sub

'Save userid & password in form for connections, and on HiddenSettings tab if checkbox = true
Private Sub cmdPwdOK_Click()
Debug.Print Me.CheckBox1.Value
    If Me.CheckBox1.Value = True Then ' save in HiddenSettings or not
         HiddenSettings.Range("user_id").Value = Me.txtUserId.Value
         HiddenSettings.Range("rpt_pwd").Value = Me.txtPassword.Value
    End If
    Debug.Print "pwdform:" & PwdForm.txtPassword.Value
    Me.Hide ' we want txtUserId&txtPassword available
    'Unload Me
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
  'Intercept/repurpose Unload if user clicks form "X" close button.
  If CloseMode = 0 Then
    Cancel = True
    Me.Hide
  End If
End Sub
