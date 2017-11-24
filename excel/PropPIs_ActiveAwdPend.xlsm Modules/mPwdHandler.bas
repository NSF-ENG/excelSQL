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
    & ";UID=" & PwdForm.txtUserId.Value _
    & ";PWD=" & PwdForm.txtPassword.Value & ";"
End Function

'Method handles password if we have one, else show password form to request one.
'checks if password works.
Sub handlePwd()
    PwdForm.txtUserId.Value = HiddenSettings.Range("user_id").Value
    If HiddenSettings.Range("user_id").Value = "" Or HiddenSettings.Range("rpt_pwd").Value = "" Then
        PwdForm.txtPassword.Value = ""
        PwdForm.CheckBox1.Value = False
        PwdForm.Show
    Else ' try the saved password
        PwdForm.txtPassword.Value = HiddenSettings.Range("rpt_pwd").Value
    End If
    
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
        If MsgBox("The reportserver userid and password are not working; please check if they have been updated and try again.", _
             vbOKCancel) <> vbOK Then End
        Call handlePwd
        End
    End If
End Sub

Sub doQuery(qt As QueryTable, SQL As String, Optional refreshFlag As Boolean = False, Optional db As String = "rptdb")
'stuff connection and command into query, call refresh, and handle errors

   With qt
        .Connection = "ODBC;" & makeConnectionString(db)
        .CommandText = SQL
        .Refresh (refreshFlag)
    End With
ExitHandler: Exit Sub
ErrHandler:
End Sub

' Temporarily stuff userId and rptPassword from PwdForm in all ODBCConnections in worksheet
'
Sub oldFixConnections(userId As String, rptPassword As String)
' called from user PwdForm code and handlePwd above; calls handlePwd again if password doesn't work
  Dim cstring As String
  Dim i, locUID, locPWD As Long
  
  #If Mac Then
    cstring = HiddenSettings.Range("Mac_connect_string")
  #Else
    cstring = HiddenSettings.Range("PC_connect_string")
  #End If
  If Right(cstring, 1) <> ";" Then cstring = cstring & ";"
  cstring = cstring & "UID=" & userId & ";PWD=" & rptPassword & ";"
   
  If Len(cstring) > 5 Then 'Open a connection briefly to make sure the password is good.
    Dim cn As Object
    Dim good As Boolean
    Set cn = CreateObject("ADODB.Connection")
    With cn
      .ConnectionString = cstring
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
        If MsgBox("The reportserver userid and password are not working; please check if they have been updated and try again.", _
             vbOKCancel) <> vbOK Then End
        Call handlePwd
        End
    End If
  End If
  Dim ws As Worksheet
  Dim lo As ListObject
  Dim qt As QueryTable
  
  On Error Resume Next ' stuff odbc connection strings
  For Each ws In ThisWorkbook.Sheets
    For Each lo In ws.ListObjects
      If Left(lo.QueryTable.Connection, 4) = "ODBC" Then lo.QueryTable.Connection = "ODBC;" & cstring
    Next
  Next
  On Error GoTo 0
  
'  For i = 1 To ThisWorkbook.Connections.count
'    With ThisWorkbook.Connections(i).ODBCConnection
'        .Connection = "ODBC;" & cstring
'    End With
'  Next
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
