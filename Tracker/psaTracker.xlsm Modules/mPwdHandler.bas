Attribute VB_Name = "mPwdHandler"
Option Explicit
' Jack Snoeyink      Oct 2, 2017
' This module depends on a PwdForm and Settings tab with input boxes txtUserID and txtPwd to manage reportserver passwords.
' If Settings does not contain a password, it will use PwdForm to request one.
' A checkbox on the form will optionally save the userid & pwd back to Settings.
' Settings should stay formatted so password shows as *******

'Method handles password if we have one, else show password form to request one.
Sub handlePwd()
    If Settings.txtUserID.Value = "" Or Settings.txtPwd.Value = "" Then
        PwdForm.txtUserID.Value = Settings.txtUserID.Value
        PwdForm.txtPassword.Value = ""
        PwdForm.CheckBox1.Value = False
        PwdForm.Show
    Else ' try the saved password
        Call FixConnections(Settings.txtUserID.Value, Settings.txtPwd.Value)
    End If
End Sub


' Temporarily stuff userId and rptPassword from PwdForm in all ODBCConnections in worksheet
'
Sub FixConnections(userId As String, rptPassword As String)
' called from user PwdForm code and handlePwd above; calls handlePwd again if password doesn't work
  Dim cstring As String
  Dim i, locUID, locPWD As Long
  For i = 1 To ThisWorkbook.Connections.count
    With ThisWorkbook.Connections(i).ODBCConnection
        cstring = .Connection
        'locUID = InStrRev(cstring, "UID=") ' need to parse to handle mac connection strings
        'locPWD = InStrRev(cstring, "PWD=")
        'locSemi = InStrRev(cstring, ";")
        cstring = Left(cstring, InStrRev(cstring, "UID=") + 3) & userId & ";PWD=" & rptPassword & ";"
        ' MsgBox cstring ' note: the password will not be saved with the connection, but is used during the session.
        .Connection = cstring
       ' MsgBox .Connection
    End With
  Next
  
  If Len(cstring) > 5 Then 'Open a connection briefly to make sure the password is good.
    Dim cn As Object
    Dim good As Boolean
    Set cn = CreateObject("ADODB.Connection")
    With cn
      .ConnectionString = Mid(cstring, 6) ' leave off initial ODBC;
      .ConnectionTimeout = 10 ' in seconds
      On Error Resume Next
      .Open
      good = Err.Number = 0 ' if any error, we couldn't open connection.
      .Close
    End With
    On Error GoTo 0
    Set cn = Nothing
    If Not good Then
        Settings.txtPwd.Value = ""
        If MsgBox("The reportserver userid and password are not working; please check if they have been updated and try again.", _
             vbOKCancel) <> vbOK Then End
        Call handlePwd
        End
    End If
  End If
End Sub



