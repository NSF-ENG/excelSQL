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
Global gLastConnectionError As String

Sub activateApp()
' activate the application window
    DoEvents
    #If Mac Then
    AppActivate Application.name
    #Else 'PC
    AppActivate Application.Caption
    #End If 'PC
    DoEvents
End Sub

Function makeConnectionString(Optional db As String = "rptdb") As String
' put database, UID, and PWD at end of Mac or PC connection string
Dim cstring As String
  #If Mac Then
    cstring = HiddenSettings.Range("Mac_connect_string")
  #Else
    cstring = HiddenSettings.Range("PC_connect_string")
  #End If
  If VBA.Right$(cstring, 1) <> ";" Then cstring = cstring & ";"
  makeConnectionString = cstring & "database=" & db _
    & ";UID=" & gPwdForm.txtUserId.Value _
    & ";PWD=" & gPwdForm.txtPassword.Value & ";"
 'Debug.Print "mc:" & makeConnectionString
End Function

Function needPassword() As Boolean
' this function checks if the form does not hold a valid password
' On a PC, it opens an ADODB connection and checks for error
' A Mac equivalent for Openlink may come from comments on this support case:
'http://support.openlinksw.com/support/techupdate.vsp?c=22287
'user: jsnoeyin@nsf.gov pwd: pwd123!
'
' In particular, I'd suggest seeing if AppleScript calls can run their test tool:
'Note that we provide 2 test tools, each in a Unicode and a non-Unicode ("ANSI") flavor, which should be used to test the corresponding driver.
'"iODBC Test.command" launches a command-line tool within a fresh Terminal.app session.
'It may also be worth testing Excel with a simplified connect string in your workbook, that relies on the DSN configuration, and so reads simply --
'   ODBC;DSN=rptServer;UID=ccfuser

   needPassword = True
   If gPwdForm Is Nothing Then Exit Function
   If gPwdForm.txtUserId.Value = "" Or gPwdForm.txtPassword.Value = "" Then Exit Function
   needPassword = tryConnection() <> 0 ' if any error, we still need a password
End Function
Function tryConnection(Optional SQL As String = "") As Long
'
   #If Mac Then
     tryConnection = True ' assume we are ok (MAC FIX)
     ' maybe use applescript to check connection & password
     ' see comment above
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
      tryConnection = Err.Number 'return error number
      gLastConnectionError = Err.Description
      If tryConnection = 0 And Len(SQL) > 0 Then
        Dim rs As Object
        Set rs = cn.Execute(SQL)
        tryConnection = Err.Number
        gLastConnectionError = Err.Description
        rs.Close
      End If
      .Close
    End With
    On Error GoTo 0
    Set rs = Nothing
    Set cn = Nothing
   #End If
End Function
Function lastConnectionErrorDescription()
'hack to get description, which comes from Sybase
  lastConnectionErrorDescription = gLastConnectionError
End Function

Function handlePwd() As Boolean
'Function returns true if it has a valid reportserver password or obtains on from the user.
'checks if password works.
If gPwdForm Is Nothing Then ' populate form from saved password
  Set gPwdForm = New PwdForm
End If
If needPassword Then
  With gPwdForm
    .CheckBox1.Value = False
    .Show  ' this returns with a good password, or the user did not want to enter one, so .txtPassword.Value = "".
  End With
  handlePwd = Not needPassword
Else
  handlePwd = True
End If
End Function

Public Sub doQuery(QT As QueryTable, SQL As String, Optional backgroundFlag As Boolean = False, Optional db As String = "rptdb")
'stuff connection and command into query, call refresh, and handle errors
' we assume that password has been tested unless we see an error.
' Note: try out queries with backgroundFlag False to catch errors.
    'Debug.Print "doQuery: " & (gPwdForm Is Nothing)
If gPwdForm Is Nothing And Not handlePwd Then Exit Sub ' abort
RetryHandler:
DoEvents
On Error GoTo errHandler
  With QT
    .Connection = "ODBC;" & makeConnectionString(db)
    #If Mac Then
    .SQL = SQL
    #Else
    .CommandText = SQL
    #End If
    .Refresh (backgroundFlag)
  End With
ExitHandler:
  On Error GoTo 0
  Exit Sub
errHandler:
    activateApp
    Select Case MsgBox("doQuery Error on " & db & " query " & VBA.Left$(SQL, 50) & vbNewLine & Err.Number & ":" & Err.Description, vbAbortRetryIgnore)
    Case vbAbort: End
    Case vbRetry: If handlePwd Then GoTo RetryHandler
    End Select
GoTo ExitHandler
End Sub


