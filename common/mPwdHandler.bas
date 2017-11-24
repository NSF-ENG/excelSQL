Attribute VB_Name = "mPwdHandler"
Option Explicit
' Jack Snoeyink      Oct 2, 2017
' This module depends on a PwdForm and ReportServerPwd tab with input boxes txtUserID and txtPwd to manage reportserver passwords.
' If ReportServerPwd does not contain a password, it will use PwdForm to request one.
' A checkbox on the form will optionally save the userid & pwd back to ReportServerPwd.
' ReportServerPwd should stay formatted so password shows as *******

'Method handles password if we have one, else show password form to request one.
Sub handlePwd()
    If HiddenSettings.Range("user_id").Value = "" Or HiddenSettings.Range("rpt_pwd").Value = "" Then
        PwdForm.txtUserId.Value = HiddenSettings.Range("user_id").Value
        PwdForm.txtPassword.Value = ""
        PwdForm.CheckBox1.Value = False
        PwdForm.Show
    Else ' try the saved password
        Call FixConnections(HiddenSettings.Range("user_id").Value, HiddenSettings.Range("rpt_pwd").Value)
    End If
End Sub

' Temporarily stuff userId and rptPassword from PwdForm in all ODBCConnections in worksheet
'
Sub FixConnections(userId As String, rptPassword As String)
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
