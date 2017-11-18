Attribute VB_Name = "Module_Pwd"
' Credentials sheet, pwdForm, Module_Pwd work together to allow password entry once for rptserver connections
' Uses ranges for user_id and rpt_pwd on protected Credentials sheet that it puts into global variables:
' This hack assumes that all connections are to the reportserver, and does not error checking/recovery

' Jack Snoeyink  Aug 2016

' Credentials sheet, pwdForm, Module_Pwd work together to allow password entry once for rptserver connections
' Uses ranges for user_id and rpt_pwd on protected Credentials sheet that it puts into global variables:
' This hack assumes that all connections are to the reportserver, and does not error checking/recovery

' Jack Snoeyink  Aug 2016

Option Explicit
Public userId As String
Public rptPassword As String

' Temporarily stuff userId and rptPassword from PwdForm in all ODBCConnections in worksheet
Sub FixConnections()
' called from user PwdForm code and handlePwdForm below
  Dim cstring As String
  Dim i As Long
  For i = 1 To ThisWorkbook.Connections.COUNT
    With ThisWorkbook.Connections(i).ODBCConnection
        cstring = .Connection
        cstring = Left(cstring, InStrRev(cstring, "UID=") + 3) & userId & ";PWD=" & rptPassword & ";"
        ' MsgBox cstring ' note: the password will not be saved with the connection, but is used during the session.
        .Connection = cstring
       ' MsgBox .Connection
    End With
  Next
End Sub

Public Sub handlePwdForm()
 With ThisWorkbook.Worksheets("Settings")
        PwdForm.txtUserID.Value = Sheet7.UserID2.Value
        userId = PwdForm.txtUserID.Value
        PwdForm.txtPassword.Value = Sheet7.pwd_3.Value
        rptPassword = PwdForm.txtPassword.Value
 End With
 
    Call FixConnections
    
'    If PwdForm.CheckBox1.Value = True Then
'        PwdForm.txtUserId.Value = userId
'        PwdForm.txtPassword.Value = rptPassword
'    Else
'        PwdForm.txtUserId.Value = ""
'        PwdForm.txtPassword.Value = ""
'
'    End If
End Sub
'Method used to show log in form
Public Sub Show()
    If Sheet7.UserID2.Value = "" Or Sheet7.pwd_3.Value = "" Then
    PwdForm.txtUserID.Value = ""
    PwdForm.txtPassword.Value = ""
    PwdForm.CheckBox1.Value = False
     PwdForm.Show
    Else
    Call handlePwdForm
    End If
End Sub



