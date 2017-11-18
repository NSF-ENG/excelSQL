Attribute VB_Name = "Module1"
Option Explicit

Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Private Declare PtrSafe Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)
Private Declare PtrSafe Function MapVirtualKey Lib "user32" Alias "MapVirtualKeyA" (ByVal wCode As Long, ByVal wMapType As Long) As Long
Private Const VK_RETURN = &HD

Sub pageInfo()
' go to a page and report forms, links, anchors, etc.
Dim IE As InternetExplorerMedium ' This object (the "medium" variety as opposed to "InternetExplorer") is necessary in our security climate
Dim i, n As Long
Dim x, currentText, searchText, oldURL As String
x = "https://www.ejacket.nsf.gov" ' Sheet1.Range("F1").Value ' change sheet accordingly

    Set IE = New InternetExplorerMedium
    IE.Navigate ("https://www.ejacket.nsf.gov")
    Call myWait(IE)
    IE.Navigate ("https://www.ejacket.nsf.gov/ej/showProposal.do?Continue=Y&optimize=Y&ID=" & 1652695)
    Call myWait(IE)
    IE.Navigate ("https://www.ejacket.nsf.gov/ej/showDataMaintenanceInfo.do?dispatch=getSpde&fromleftNav=Y")
    Call myWait(IE)
    IE.Visible = True
    n = IE.Document.forms.Length
    Debug.Print n & " forms"
    For i = 0 To n - 1
        Debug.Print i & ":" & IE.Document.forms(i).Name & "," & IE.Document.forms(i).InnerText
    Next
    
    n = IE.Document.Links.Length
    Debug.Print n & " links"
    For i = 0 To n - 1
        Debug.Print i & ":" & IE.Document.Links(i).Name & "," & IE.Document.Links(i).InnerText
    Next

    n = IE.Document.anchors.Length
    Debug.Print n & " anchors"
    For i = 0 To n - 1
        Debug.Print i & ":" & IE.Document.anchors(i).Name & "," & IE.Document.anchors(i).InnerText
    Next
    
    n = IE.Document.all.Length
    Debug.Print n & " all"
    For i = 0 To n - 1
        Debug.Print i & ":" & IE.Document.all(i).Name & "," & IE.Document.all(i).InnerText
    Next
End Sub

Private Sub myWait(IE)
    ' wait for IE to be ready
    Dim count As Long
    Sleep 10 * Range("delayTime").Value
    count = 0
    While IE.Busy And (Not IE.ReadyState = READYSTATE_COMPLETE) And (count < 40)
        DoEvents
        Sleep 10 * (Range("delayTime").Value + count)
        DoEvents
        count = count + 1
    Wend
    If count > 35 Then MsgBox count & " in myWait.  We seem to be having problems."
End Sub

Private Sub CheckCollabs(IE)
    ' cell in spreadsheet should be Y,Yes, or yes if we want to apply to collabs.
    If (UCase(Left(Range("apply2Collabs").Value, 1)) = "Y") Then
    ' look for apply to collabs button on IE and check it if it is present
      Call myWait(IE)
      If (IE.Document.getElementsByName("applyToCollabs").Length > 0) Then
       IE.Document.getElementsByName("applyToCollabs").Item(0).Click
        Call myWait(IE)
      End If
    End If
End Sub


' use userId and rptPassword from Credentials sheet, if present
Public Sub FixConnections()
    Dim i As Long
    Dim userId, pwdString, cstring As String
    
    userId = Trim(Range("user_id").Value)
    If Len(userId) < 2 Then ' if no user, stop and make them fill in credentials
      MsgBox "Cannot query until you enter a reportServer userid on the Credentials tab"
      End
      End If
      pwdString = ";PWD=" & Trim(Range("rpt_pwd").Value)
      If Len(pwdString) < 6 Then pwdString = "" ' if no password, user will supply
      
      For i = 1 To ThisWorkbook.Connections.count
        With ThisWorkbook.Connections(i).ODBCConnection
            cstring = .Connection
            cstring = Left(cstring, InStrRev(cstring, "UID=") + 3) & userId & pwdString & ";"
            ' MsgBox cstring ' note: the password will not be saved with the connection, but is used during the session.
            .Connection = cstring
           ' MsgBox .Connection
        End With
      Next
End Sub

Sub checkAutocoding() ' Query to see if autocoded data is there
Dim tbl As Range
Dim i As Long
Dim prop_id As String

Call FixConnections

Set tbl = Range("propTable[prop_id]")
If tbl.Rows.count > 1 Then  ' list the jackets to be coded
 prop_id = "('" & Join(Application.Transpose(tbl.Value), "','") & "')"
 prop_id = " in " & Replace(Replace(Replace(prop_id, " ", ""), Chr(160), ""), ",''", "")
Else
 prop_id = " = '" & Trim(Replace(tbl.Value, Chr(160), "")) & "'"
End If

Call tbl.AutoFilter ' reset filters
Call tbl.AutoFilter
Call Range("CheckPRCs[prop_id]").AutoFilter
Call Range("CheckPRCs[prop_id]").AutoFilter

Dim QT As QueryTable
With Worksheets("AutoCode").ListObjects
    For i = 1 To .count
        If .Item(i).Name = "checkPRCs" Then Set QT = .Item(i).QueryTable
    Next i
End With

With QT
    .CommandText = "SET NOCOUNT ON" & vbLf _
        & "SELECT prop.prop_id INTO #myProps FROM csd.prop prop WHERE " & vbLf _
        & "(prop.prop_id " & prop_id & " OR prop.lead_prop_id " & prop_id & ")" & vbLf _
        & "SELECT DISTINCT prop.prop_id, pa.prop_atr_code, id=identity(18), 0 as 'seq' INTO #myPRCs" & vbLf _
        & "FROM #myProps prop, csd.prop_atr pa WHERE pa.prop_id = prop.prop_id  AND pa.prop_atr_type_code = 'PRC'" & vbLf _
        & "ORDER BY prop.prop_id, pa.prop_atr_code" & vbLf _
        & "SELECT prop_id, MIN(id) as 'start' INTO #myStarts FROM #myPRCs GROUP BY prop_id" & vbLf _
        & "UPDATE #myPRCs set seq = id-M.start FROM #myPRCs r, #myStarts M WHERE r.prop_id = M.prop_id" & vbLf _
        & "SELECT prop.prop_id, prop.cntx_stmt_id, prop.bas_rsch_pct, prop_stts.prop_stts_abbr, prop.pm_ibm_logn_id, pi_vw.pi_last_name, pi_vw.pi_frst_name, inst.inst_name," & vbLf _
        & "prop.org_code, CASE WHEN prop.org_code<>prop.orig_org_code THEN prop.orig_org_code END AS orig_org," & vbLf _
        & "prop.pgm_ele_code, CASE WHEN prop.pgm_ele_code <> prop.orig_pgm_ele_code THEN prop.orig_pgm_ele_code END AS orig_PEC, " & vbLf _
        & "(SELECT MAX( CASE pa.seq WHEN 0 THEN pa.prop_atr_code ELSE '' END ) + ' ' +" & vbLf _
        & "        MAX( CASE pa.seq WHEN 1 THEN pa.prop_atr_code ELSE '' END ) + ' ' +" & vbLf _
        & "        MAX( CASE pa.seq WHEN 2 THEN pa.prop_atr_code ELSE '' END ) + ' ' +" & vbLf _
        & "        MAX( CASE pa.seq WHEN 3 THEN pa.prop_atr_code ELSE '' END ) + ' ' +" & vbLf _
        & "        MAX( CASE pa.seq WHEN 4 THEN pa.prop_atr_code ELSE '' END ) + ' ' +" & vbLf _
        & "        MAX( CASE pa.seq WHEN 5 THEN pa.prop_atr_code ELSE '' END ) + ' ' +" & vbLf _
        & "        MAX( CASE pa.seq WHEN 6 THEN pa.prop_atr_code ELSE '' END ) FROM #myPRCs pa WHERE p.prop_id = pa.prop_id) AS PRCs" & vbLf _
        & "FROM #myProps p, csd.inst inst, csd.pi_vw pi_vw, csd.prop prop, csd.prop_stts prop_stts" & vbLf _
        & "WHERE prop.pi_id = pi_vw.pi_id AND prop.inst_id = inst.inst_id AND p.prop_id = prop.prop_id AND prop.prop_stts_code = prop_stts.prop_stts_code" & vbLf _
        & "DROP TABLE #myStarts DROP TABLE #myPRCs DROP TABLE #myProps" & vbLf
    .Refresh BackgroundQuery:=False
End With

End Sub

Sub autoCode() ' autocode based on propTable: set research to 100%, context statement, & PRCs
Dim IE As InternetExplorerMedium ' This object (the "medium" variety as opposed to "InternetExplorer") is necessary in our security climate
Dim tbl As Range
Dim i, j, r, c As Long
Dim prop_id, ctxt, prc As String
Dim havePRC As Boolean
        
Set tbl = Range("propTable[[prop_id]:[prc3]]") ' prop_id is col 1, ctxt 2, prcs 3 - c
r = tbl.Rows.count
c = tbl.Columns.count

Set IE = New InternetExplorerMedium
IE.Navigate ("https://www.ejacket.nsf.gov")
Call myWait(IE)
IE.Visible = True
DoEvents
AppActivate Application.Caption
DoEvents
If MsgBox("Be sure that Internet Explorer is logged in to eJacket." & vbLf & "   Ready to autocode " & r & " jackets, " _
          & tbl(1, 1).Value & " to " & tbl(r, 1).Value & "?", vbOKCancel) <> vbOK Then End ' show text gotten, allow cancel
IE.Visible = False
DoEvents
For i = 1 To r
    prop_id = Trim(Replace(tbl(i, 1).Value, Chr(160), ""))  ' column 1 is prop_id
    If (Len(prop_id) = 7) Then ' Probably have a prop_id; go to Jacket
        IE.Navigate ("https://www.ejacket.nsf.gov/ej/showProposal.do?Continue=Y&optimize=Y&ID=" & prop_id)
        Call myWait(IE)
        IE.Visible = True
        Call myWait(IE)
        
        If (UCase(Left(Range("basicRsrch").Value, 1)) = "Y") Then
        ' set research % to 100%, if total is not already 100%
         IE.Navigate ("https://www.ejacket.nsf.gov/ej/showDataMaintenanceInfo.do?dispatch=getRdAllotment&fromleftNav=Y")
         Call myWait(IE)
         If IE.Document.getElementsByName("rdAllotmentTotal")(0).Value = "0%" Then
             j = 0
             Do
                 With IE.Document.getElementsByName("rdAllotment.basicResearchAsPct")(0)
                     .Focus
                     .Value = "100"
                     .FireEvent ("onblur")
                 End With
                 Call myWait(IE)
                 j = j + 1 ' try three times then give up
             Loop While IE.Document.getElementsByName("rdAllotmentTotal")(0).Value = "0%" And j < 3
             Call CheckCollabs(IE)
             IE.Document.forms("dataMaintenanceForm").submit
             Call myWait(IE)
         End If
        End If

        ' set context statement
        ctxt = Trim(tbl(i, 2).Value) ' col 2 is ctxt
        If (UCase(Left(Range("assignCtxIndividually").Value, 1)) = "Y") And Len(ctxt) > 0 Then ' assign context
            IE.Navigate ("https://www.ejacket.nsf.gov/ej/showDataMaintenanceInfo.do?dispatch=getReviewSummary&fromleftNav=Y")
            Call myWait(IE)
            With IE.Document.getElementsByName("reviewSummary.contextStatementId")(0)
                .Focus
                .Value = ctxt
            End With
            Call myWait(IE)
            Call CheckCollabs(IE)
            IE.Document.forms("dataMaintenanceForm").submit
            Call myWait(IE)
        End If

        ' set PRCs if any
        havePRC = False
        For j = 3 To c  ' col 3-c are prcs
            prc = Trim(tbl(i, j).Value)
            If Len(prc) > 0 Then
                If Not (havePRC) Then 'have first PRC: go to PRC page
                    IE.Navigate ("https://www.ejacket.nsf.gov/ej/showDataMaintenanceInfo.do?dispatch=getSpde&fromleftNav=Y")
                    Call myWait(IE)
                End If
                havePRC = True

                With IE.Document.getElementsByName("newPgmRefCode")(0) ' add new PRC
                    .Focus
                    .Value = prc
                End With
                Call myWait(IE)
                IE.Document.getElementsByName("addPgmRef")(0).Click
                Call myWait(IE)

            End If
        Next j
        If havePRC Then ' save PRCs
            Call CheckCollabs(IE)
            IE.Document.forms("dataMaintenanceForm").submit
            Call myWait(IE)
        End If
    End If
Next i

IE.Quit
Set IE = Nothing
DoEvents
Sleep 2000
DoEvents
Call checkAutocoding
End Sub

Sub associateCtxt()
Dim IE As InternetExplorerMedium ' This object (the "medium" variety as opposed to "InternetExplorer") is necessary in our security climate
Dim tbl As Range
Dim i, j, k, r, n As Integer
Dim prop_id, ctxtId As String
        
Set tbl = Range("propTable[prop_id]")
r = tbl.Rows.count
ctxtId = Trim(Range("context_id").Value)

Set IE = New InternetExplorerMedium
IE.Navigate ("https://www.ejacket.nsf.gov/ej/searchContextStatements.do?dispatch=show")
IE.Visible = True

ThisWorkbook.Activate
If Len(ctxtId) < 1 Then
   MsgBox "Search for your Context Statement Id and enter it on the spreadsheet, then try again"
   End
End If
If MsgBox("Be sure that Internet Explorer is logged in to eJacket." & vbLf & "   Ready to associate " & ctxtId & " to " & r & " jackets, " _
          & tbl(1, 1).Value & " to " & tbl(r, 1).Value & "?", vbOKCancel) <> vbOK Then End ' show text gotten, allow cancel

IE.Visible = False
DoEvents
IE.Navigate ("https://www.ejacket.nsf.gov/ej/showContextStatementDetail.do?contextID=" & ctxtId)
Call myWait(IE)
IE.Navigate ("https://www.ejacket.nsf.gov/ej/processAssociateContextStatement.do?dispatch=showLookup&contextID=" & ctxtId)
IE.Visible = True
Call myWait(IE)

j = 0
For i = 1 To r
    If j = 0 Then
        n = IE.Document.Links.Length
        For k = 0 To n - 1 ' find link to associate by proposal ID(s)
            If (IE.Document.Links(k).InnerText = "By Proposal ID(s)") Then
                IE.Document.Links(k).Click
                Call myWait(IE)
                k = n
            End If
        Next k
    End If
    prop_id = Trim(Replace(tbl(i).Value, Chr(160), ""))  ' first column is prop_id
    If (Len(prop_id) = 7) Then ' Probably have a prop_id; go to Jacket
        IE.Document.getElementsByName("proposalIds")(j).Value = prop_id
        Call myWait(IE)
        j = j + 1
    End If
    If (j = 10) Or (i = r) Then ' last on page
        IE.Document.getElementsByName("associateButton")(0).Click
        Call myWait(IE)
        j = 0
        IE.Navigate ("https://www.ejacket.nsf.gov/ej/processAssociateContextStatement.do?dispatch=showLookup&contextID=" & ctxtId)
        Call myWait(IE)
    End If
Next i

IE.Quit
Set IE = Nothing
DoEvents
Sleep 2000
DoEvents
Call checkAutocoding
End Sub
