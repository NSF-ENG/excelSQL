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
Global gPwdForm As PwdForm

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
 'Debug.Print "mc:" & makeConnectionString
End Function

Function needPassword() As Boolean
   needPassword = True
   If gPwdForm Is Nothing Then Exit Function
   If gPwdForm.txtPassword.Value = "" Then Exit Function
   #If Mac Then
   needPassword = False ' assume we are ok (FIX)
   ' maybe use applescript to check connection & password
   #Else
' use ADODB connection to try password; get a fresh one if it has expired.
' Need to check that this actually uses the password
    Dim cn As Object
    Set cn = CreateObject("ADODB.Connection")
    With cn
      .ConnectionString = makeConnectionString
      .ConnectionTimeout = 10 ' in seconds
      On Error Resume Next
      .Open
      needPassword = Err.Number <> 0 ' if any error, we couldn't open connection.
      .Close
    End With
    On Error GoTo 0
    Set cn = Nothing

   #End If
End Function

'Method handles password if we have one, else show password form to request one.
'checks if password works.
Sub handlePwd()
    'Debug.Print "handle: " & (gPwdForm Is Nothing)
    If gPwdForm Is Nothing Then Set gPwdForm = New PwdForm
    With gPwdForm
    .txtUserId.Value = HiddenSettings.Range("user_id").Value
    'Debug.Print "pwdval:" & .txtUserID.Value
    If HiddenSettings.Range("user_id").Value = "" Or HiddenSettings.Range("rpt_pwd").Value = "" Then
        .txtPassword.Value = ""
        .CheckBox1.Value = False
        .Show
    Else ' try the saved password
        .txtPassword.Value = HiddenSettings.Range("rpt_pwd").Value
    End If
    'Debug.Print "pwdout:" & .txtPassword.Value
    End With
    #If Mac Then
    ' need to supply
    #Else
    If needPassword Then
        HiddenSettings.Range("rpt_pwd").Value = ""
        AppActivate Application.Caption
        DoEvents
        If MsgBox("The reportserver userid and password are not working; please check if they have been updated and try again." _
        & vbNewLine & "If remote, ensure you have an active VPN connection into the NSF network.", vbOKCancel) <> vbOK Then End
        Call handlePwd
    End If
    #End If
End Sub

Public Sub doQuery(qt As QueryTable, SQL As String, Optional backgroundFlag As Boolean = False, Optional db As String = "rptdb")
'stuff connection and command into query, call refresh, and handle errors
' Note: try out queries with backgroundFlag False to catch errors.
    
    'Debug.Print "doQuery: " & (gPwdForm Is Nothing)
   If gPwdForm Is Nothing Then Call handlePwd
RetryHandler:
On Error GoTo ErrHandler
   With qt
        .Connection = "ODBC;" & makeConnectionString(db)
        #If Mac Then
        .SQL = SQL
        #Else
        .CommandText = SQL
        #End If
        .Refresh (backgroundFlag)
    End With
ExitHandler: Exit Sub
ErrHandler:
    Dim rtn As Integer
    rtn = MsgBox("doQuery Error " & Err.Number & ":" & Err.Description, vbAbortRetryIgnore)
    If rtn = vbAbort Then End
    If rtn = vbRetry Then GoTo RetryHandler
    GoTo ExitHandler
End Sub


