Attribute VB_Name = "mPwdHandler"
Option Explicit
' Jack Snoeyink      Oct 2, 2017
' This module depends on a PwdForm with input boxes txtUserID and txtPwd
' and a HiddenSettings tab with cells labeled 'user_id' and 'rpt_pwd'
' to manage reportserver passwords.
' tryPwd uses an ADODB connection to try th epassword in PwdForm
' If HiddenSettings does not contain a password, it will use PwdForm to request one.
' A checkbox on the form will optionally save the userid & pwd back to HiddenSettings.
' HiddenSettings should stay formatted so password shows as *******

Function makeConnectionString(Optional db As String = "rptdb") As String
' put database, UID, and PWD at end of Mac or PC connection string
Dim cstring As String
  #If Mac Then
    cstring = HiddenSettings.Range("Mac_connect_string")
  #Else
    cstring = HiddenSettings.Range("PC_connect_string")
  #End If
  If Right(cstring, 1) <> ";" Then cstring = cstring & ";"
  makeConnectionString = cstring & "database=" & db _
    & ";UID=" & gPwdForm.txtUserId.Value _
    & ";PWD=" & gPwdForm.txtPassword.Value & ";"
End Function

'Method handles password if we have one, else show password form to request one.
'checks if password works.
Sub handlePwd()
    If gPwdForm Is Nothing Then Set gPwdForm = New PwdForm
    With gPwdForm
    .txtUserId.Value = HiddenSettings.Range("user_id").Value
    If HiddenSettings.Range("user_id").Value = "" Or HiddenSettings.Range("rpt_pwd").Value = "" Then
        .txtPassword.Value = ""
        .CheckBox1.Value = False
        .Show
    Else ' try the saved password
        .txtPassword.Value = HiddenSettings.Range("rpt_pwd").Value
    End If
    End With
    #If Mac Then
    #Else
' use ADODB connection to try password; get a fresh one if it has expired.
    Dim cn As Object
    Dim good As Boolean
    Set cn = CreateObject("ADODB.Connection")
    With cn
      .ConnectionString = makeConnectionString
      .ConnectionTimeout = 10 ' in seconds
      On Error Resume Next
      .Open
      good = Err.Number = 0 ' if any error, we couldn't open connection.
      .Close
    End With
    On Error GoTo 0
    Set cn = Nothing
    If Not good Then
        HiddenSettings.Range("rpt_pwd").Value = ""
        If MsgBox("The reportserver userid and password are not working; please check if they have been updated and try again." _
                  & vbNewLine & "If remote, ensure you have an active VPN connection into the NSF network.", vbOKCancel) <> vbOK Then End
        Call handlePwd
        End
    End If
    #End If
End Sub

Public Sub doQuery(qt As QueryTable, SQL As String, Optional refreshFlag As Boolean = False, Optional db As String = "rptdb")
'stuff connection and command into query, call refresh, and handle errors
   If gPwdForm Is Nothing Then Call handlePwd
   
   With qt
        .Connection = "ODBC;" & makeConnectionString(db)
        #If Mac Then
        .SQL = SQL
        #Else
        .CommandText = SQL
        #End If
        .Refresh (refreshFlag)
    End With
ExitHandler: Exit Sub
ErrHandler:
End Sub

