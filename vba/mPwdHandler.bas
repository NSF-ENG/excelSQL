Attribute VB_Name = "mPwdHandler"
Option Explicit
' Jack Snoeyink      Oct 2, 2017
' This module depends on a PwdForm and ReportServerPwd tab with input boxes txtUserID and txtPwd to manage reportserver passwords.
' If ReportServerPwd does not contain a password, it will use PwdForm to request one.
' A checkbox on the form will optionally save the userid & pwd back to ReportServerPwd.
' ReportServerPwd should stay formatted so password shows as *******

'Method handles password if we have one, else show password form to request one.
Sub handlePwd()
    If ReportServerPwd.txtUserId.Value = "" Or ReportServerPwd.txtPwd.Value = "" Then
        PwdForm.txtUserId.Value = ReportServerPwd.txtUserId.Value
        PwdForm.txtPassword.Value = ""
        PwdForm.CheckBox1.Value = False
        PwdForm.Show
    Else ' try the saved password
        Call FixConnections(ReportServerPwd.txtUserId.Value, ReportServerPwd.txtPwd.Value)
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
        cstring = stuffParam(stuffParam(.Connection, "UID=", userId), "PWD=", rptPassword)
        ' MsgBox cstring ' note: the password will not be saved with the connection, but is used during the session.
        .Connection = cstring
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
        ReportServerPwd.txtPwd.Value = ""
        If MsgBox("The reportserver userid and password are not working; please check if they have been updated and try again.", _
             vbOKCancel) <> vbOK Then End
        Call handlePwd
        End
    End If
  End If
End Sub



Private Function stuffParam(s As String, paramName As String, paramValue As String) As String
'Look for paramName (eg. UID= in a connection string) then assume the value is terminated by ; or the end of string
'Replace with paramName & paramValue, or append if param was not found.
Dim locStart, locEnd As Long
    locStart = InStr(s, paramName)
    If locStart = 0 Then
       stuffParam = s & ";" & paramName & paramValue
    Else
       locEnd = InStr(locStart + Len(paramName), s, ";")
       If locEnd = 0 Then locEnd = Len(s) + 1
       stuffParam = Left$(s, locStart - 1) & paramName & paramValue & Mid$(s, locEnd)
    End If
End Function

Private Sub test_stuffParam()
Debug.Print stuffParam("UID=", "ccfuser", ";UID=abcdef;")
Debug.Print stuffParam("UID=", "ccfuser", "UID=abcdef;")
Debug.Print stuffParam("UID=", "ccfuser", "") ' this adds an extra semicolon
Debug.Print stuffParam("UID=", "ccfuser", "junk")
Debug.Print stuffParam("UID=", "ccfuser", "Test;UID=abcdef")
Debug.Print stuffParam("UID=", "ccfuser", "TestUID=abcdef; more")
End Sub
