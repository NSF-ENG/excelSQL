VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsPwd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Settings sheet, pwdForm, & cModule_Pwd work together to allow password entry once for rptserver connections
' Uses ranges for user_id and rpt_pwd on protected Settings sheet that it puts into global variables:
' In a Global variables module, include
' Global pwdHandler As New clsPwd

' Jack Snoeyink  Aug 2016

Option Explicit
Dim userId As String
Dim rptPassword As String

Public Property Let User_Id(uid As String)
' set report server userId
  userId = uid
End Property

Public Property Let Rpt_Password(pwd As String)
' set report server password
  rptPassword = pwd
End Property

Public Sub clearPwd()
' clear user id and password
  userId = ""
  rptPassword = ""
End Sub

' Temporarily stuff userId and rptPassword from PwdForm in all ODBCConnections in worksheet
Sub FixConnections()
' called from user PwdForm code and handlePwdForm below
  Dim cstring As String
  Dim i As Long
  For i = 1 To ThisWorkbook.Connections.Count
   With ThisWorkbook.Connections(i)
      If .Type = xlConnectionTypeODBC Then
       With .ODBCConnection
        cstring = .Connection
        cstring = Left(cstring, InStrRev(cstring, "UID=") + 3) & userId & ";PWD=" & rptPassword & ";"
        ' MsgBox cstring ' note: the password will not be saved with the connection, but is used during the session.
        .Connection = cstring
       ' MsgBox .Connection
       End With
     End If
    End With
  Next
End Sub

Public Sub doPwd()
' call this in your code to handle Settings once
  With ThisWorkbook.Worksheets("Settings")
    If Len(.Range("user_id").Value) > 0 Then ' if there is a user set, use it
       userId = .Range("user_id").Value
       PwdForm.txtUserId.Value = userId
    End If
     
    If Len(Trim(rptPassword)) < 1 Then
        If Len(.Range("rpt_pwd").Value) > 0 Then ' if there is a password in the sheet, use it.
          rptPassword = .Range("rpt_pwd").Value
          Call FixConnections
        Else
          PwdForm.Show
        End If
    End If
  End With
'MsgBox rptPassword
End Sub

