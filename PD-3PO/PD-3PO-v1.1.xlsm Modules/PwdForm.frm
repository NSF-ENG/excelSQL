VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} PwdForm 
   Caption         =   "Report Server Id+Password"
   ClientHeight    =   2385
   ClientLeft      =   48
   ClientTop       =   376
   ClientWidth     =   4704
   OleObjectBlob   =   "PwdForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "PwdForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdPwdCancel_Click()
  Unload Me
  End ' Abort calling program
End Sub

Private Sub cmdPwdOK_Click()
 pwdHandler.User_Id = Me.txtUserId.Value
 pwdHandler.Rpt_Password = Me.txtPassword.Value
 Call pwdHandler.FixConnections
 'Unload the form
 Unload Me
End Sub
