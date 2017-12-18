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
' This userform will be created into a global variable on the first query,
' and will hold the reportserver userid and password while the spreadsheet is open.
' We therefore need to simply hide, rather than unload.

Private Sub UserForm_Initialize()
' take saved values first time this is opened
    Me.txtUserId.Value = HiddenSettings.Range("user_id").Value
    Me.txtPassword.Value = HiddenSettings.Range("rpt_pwd").Value
End Sub

Private Sub cmdPwdCancel_Click()
  Me.Hide
  'Unload Me
  End ' abort calling program
End Sub

'Save userid & password in form for connections, and on HiddenSettings tab if checkbox = true
Private Sub cmdPwdOK_Click()
Dim rtn As Integer
    If needPassword Then ' we have a bad password.  Keep form open (retry), abort, or use "remember me" and continue (ignore)
      Select Case MsgBox("The reportserver credentials entered are not working; please check your connection to insideNSF, or whether the password was updated.", vbAbortRetryIgnore)
        Case vbAbort:
          Me.Hide
          End
        Case vbIgnore:
          If Me.CheckBox1.Value = True Then ' save in HiddenSettings
            HiddenSettings.Range("user_id").Value = Me.txtUserId.Value
            HiddenSettings.Range("rpt_pwd").Value = Me.txtPassword.Value
          End If
          Me.Hide
      End Select
    Else
        If Me.CheckBox1.Value = True Then ' save in HiddenSettings
          HiddenSettings.Range("user_id").Value = Me.txtUserId.Value
          HiddenSettings.Range("rpt_pwd").Value = Me.txtPassword.Value
        End If
        Me.Hide
    End If
End Sub


Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
  'Intercept/repurpose Unload if user clicks form "X" close button.
  If CloseMode = 0 Then
    Cancel = True
    Me.Hide
  End If
End Sub
