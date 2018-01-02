Attribute VB_Name = "mAutocoder"
Option Explicit

#If Mac Then
 #If MAC_OFFICE_VERSION >= 15 Then
  #If VBA7 Then ' 64-bit Excel 2016 for Mac
   Declare PtrSafe Function GetTickCount Lib _
"/Applications/Microsoft Excel.app/Contents/Frameworks/MicrosoftOffice.framework/MicrosoftOffice" () As Long
  #Else ' 32-bit Excel 2016 for Mac
   Declare Function GetTickCount Lib _
"/Applications/Microsoft Excel.app/Contents/Frameworks/MicrosoftOffice.framework/MicrosoftOffice" () As Long
  #End If
 #Else
  #If VBA7 Then ' does not exist, but why take a chance
   Declare PtrSafe Function GetTickCount Lib _
"Applications:Microsoft Office 2011:Office:MicrosoftOffice.framework:MicrosoftOffice" () As Long
  #Else ' 32-bit Excel 2011 for Mac
   Declare Function GetTickCount Lib _
"Applications:Microsoft Office 2011:Office:MicrosoftOffice.framework:MicrosoftOffice" () As Long
  #End If
 #End If
#Else 'PC
    #If VBA7 Then
      Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As LongPtr)
    #Else
      Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
    #End If
#End If

Private Sub BusyWait(t As Long)
#If Mac Then
' busy wait t clock ticks
    Dim EndTick As Long
    EndTick = GetTickCount + t
    Do
        DoEvents
    Loop Until GetTickCount >= EndTick
#Else 'PC
    DoEvents
    Sleep (t)
    DoEvents
#End If
End Sub

#If Mac Then ' mac does not have the HTML object libraries
#Else 'PC

Function autoPasteRA(IE As InternetExplorerMedium, prop_id As String, RA As String) As String
' stuff RA into text box using mAutocoder functions
Dim i As Integer, j As Integer
Dim overwriteQ As Variant

overwriteQ = Prefs.Range("overwrite_option").Value
If (Len(prop_id) <> 7) Then ' warn that this is not a proposal id
    autoPasteRA = prop_id & " not a prop_id" & vbNewLine
    Exit Function
End If

IE.Navigate ("https://www.ejacket.nsf.gov/ej/showProposal.do?Continue=Y&ID=" & prop_id)
Call myWait(IE)
IE.Navigate ("https://www.ejacket.nsf.gov/ej/processReviewAnalysis.do?dispatch=add&uniqId=" & prop_id & VBA.LCase$(VBA.Left$(VBA.Environ$("USERNAME"), 7)))
Call myWait(IE)

If IE.Document.getElementsByName("text")(0) Is Nothing Then
  autoPasteRA = prop_id & " can't visit eJ RA" & vbNewLine
  Exit Function
End If

With IE.Document.getElementsByName("text")(0)
  .Focus
  If (Len(.Value) < 10) Or (overwriteQ = 3) Then
   .Focus
   .Value = RA
  ElseIf (overwriteQ = 2) Then ' ask permission to overwrite
    activateApp
    If (MsgBox("OK to overwrite existing RA for " & prop_id & vbNewLine & .Value, vbOKCancel) = vbOK) Then
     .Focus
     .Value = RA
    Else ' permission not granted
      autoPasteRA = prop_id & " not overwritten." & vbNewLine
      Exit Function
    End If
  Else ' never overwrite
    autoPasteRA = prop_id & " has text in RA field." & vbNewLine
    Exit Function
  End If
End With

Call myWait(IE)
If Not IE.Document.getElementsByName("save")(0) Is Nothing Then
  IE.Document.getElementsByName("save")(0).Click
  Call myWait(IE)
  autoPasteRA = ""
Else
  autoPasteRA = prop_id & " can't save eJ RA" & vbNewLine
End If
End Function

Function openEJacket() As InternetExplorerMedium
    Set openEJacket = New InternetExplorerMedium
    openEJacket.Navigate ("https://www.ejacket.nsf.gov")
    Call myWait(openEJacket)
    openEJacket.Visible = True
End Function

Sub closeEJacket(IE)
IE.Quit
Set IE = Nothing
'Sleep 1000
'Call checkAutocoding
End Sub


Sub myWait(IE)
    ' wait for IE to be ready
    Dim count As Long
    Dim delaytime As Long
    delaytime = 10 * Prefs.Range("delayTime").Value
    BusyWait delaytime
    count = 0
    While IE.Busy And (Not IE.ReadyState = READYSTATE_COMPLETE) And (count < 40)
        BusyWait delaytime
        count = count + 1
    Wend
    If count > 35 Then
        activateApp
        MsgBox count & " in myWait.  We seem to be having problems with Internet Explorer."
    End If
End Sub

'Private Sub CheckCollabs(IE)
'    ' cell in spreadsheet should be Y,Yes, or yes if we want to apply to collabs.
'    If (UCase(vba.left$(Range("apply2Collabs").Value, 1)) = "Y") Then
'    ' look for apply to collabs button on IE and check it if it is present
'      Call myWait(IE)
'      If (IE.Document.getElementsByName("applyToCollabs").Length > 0) Then
'        IE.Document.getElementsByName("applyToCollabs").Click
'        Call myWait(IE)
'      End If
'    End If
'End Sub
'
'    IE.Navigate ("https://www.ejacket.nsf.gov/ej/showProposal.do?Continue=Y&optimize=Y&ID=" & prop_id)
'    Call myWait(IE)
'
'     IE.Navigate ("https://www.ejacket.nsf.gov/ej/showDataMaintenanceInfo.do?dispatch=getRdAllotment&fromleftNav=Y")
'     Call myWait(IE)
'     If IE.Document.getElementsByName("rdAllotmentTotal")(0).Value = "0%" Then
'         j = 0
'         Do
'             With IE.Document.getElementsByName("rdAllotment.basicResearchAsPct")(0)
'                 .Focus
'                 .Value = "100"
'                 .FireEvent ("onblur")
'             End With
'             Call myWait(IE)
'             j = j + 1 ' try three times then give up
'         Loop While IE.Document.getElementsByName("rdAllotmentTotal")(0).Value = "0%" And j < 3
'         Call CheckCollabs(IE)
'         IE.Document.forms("dataMaintenanceForm").submit
'         Call myWait(IE)
'     End If
'    End If
'
'    ' set context statement
'    ctxt = Trim(tbl(i, 2).Value)
'    If Len(ctxt) > 0 Then ' assign context
'        IE.Navigate ("https://www.ejacket.nsf.gov/ej/showDataMaintenanceInfo.do?dispatch=getReviewSummary&fromleftNav=Y")
'        Call myWait(IE)
'        With IE.Document.getElementsByName("reviewSummary.contextStatementId")(0)
'            .Focus
'            .Value = ctxt
'        End With
'        Call myWait(IE)
'        Call CheckCollabs(IE)
'        IE.Document.forms("dataMaintenanceForm").submit
'        Call myWait(IE)
'    End If
'
'        ' set PRCs if any
'        havePRC = False
'        For j = 3 To 5
'            prc = Trim(tbl(i, j).Value)
'            If Len(prc) > 0 Then
'                If Not (havePRC) Then 'have first PRC: go to PRC page
'                    IE.Navigate ("https://www.ejacket.nsf.gov/ej/showDataMaintenanceInfo.do?dispatch=getSpde&fromleftNav=Y")
'                    Call myWait(IE)
'                End If
'                havePRC = True
'
'                With IE.Document.getElementsByName("newPgmRefCode")(0) ' add new PRC
'                    .Focus
'                    .Value = prc
'                End With
'                Call myWait(IE)
'                IE.Document.getElementsByName("addPgmRef")(0).Click
'                Call myWait(IE)
'
'            End If
'        Next j
'        If havePRC Then ' save PRCs
'            Call CheckCollabs(IE)
'            IE.Document.forms("dataMaintenanceForm").submit
'            Call myWait(IE)
'        End If
'    End If
'Next i
'End Sub
'
'Sub associateCtxt()
'Dim IE As InternetExplorerMedium ' This object (the "medium" variety as opposed to "InternetExplorer") is necessary in our security climate
'Dim tbl As Range
'Dim i, j, k, r, n As Integer
'Dim prop_id As String, ctxtId As String
'
'Set tbl = Range("propIds2Context[prop_id]")
'r = tbl.Rows.count
'ctxtId = Trim(Range("context_id").Value)
'
'If (Len(ctxtId) > 0) And (MsgBox("Ready to associate " & ctxtId & " to " & r & " proposals?", vbOKCancel) <> vbOK) Then End ' allow cancel
'
'Set IE = New InternetExplorerMedium
'If Len(ctxtId) < 1 Then
'   IE.Navigate ("https://www.ejacket.nsf.gov/ej/searchContextStatements.do?dispatch=show")
'   MsgBox "Search for your Context Statement Id and enter it on the spreadsheet, then try again"
'   IE.Visible = True
'   End
'End If
'IE.Navigate ("https://www.ejacket.nsf.gov/ej/showContextStatementDetail.do?contextID=" & ctxtId)
'Call myWait(IE)
'IE.Navigate ("https://www.ejacket.nsf.gov/ej/processAssociateContextStatement.do?dispatch=showLookup&contextID=" & ctxtId)
'IE.Visible = True
'Call myWait(IE)
'
'j = 0
'For i = 1 To r
'    If j = 0 Then
'        n = IE.Document.Links.Length
'        For k = 0 To n - 1 ' find link to associate by proposal ID(s)
'            If (IE.Document.Links(k).InnerText = "By Proposal ID(s)") Then
'                IE.Document.Links(k).Click
'                Call myWait(IE)
'                k = n
'            End If
'        Next k
'    End If
'    prop_id = Trim(tbl(i).Value) ' first column is prop_id
'    If (Len(prop_id) = 7) Then ' Probably have a prop_id; go to Jacket
'        IE.Document.getElementsByName("proposalIds")(j).Value = prop_id
'        Call myWait(IE)
'        j = j + 1
'    End If
'    If (j = 10) Or (i = r) Then ' last on page
'        IE.Document.getElementsByName("associateButton")(0).Click
'        Call myWait(IE)
'        j = 0
'        IE.Navigate ("https://www.ejacket.nsf.gov/ej/processAssociateContextStatement.do?dispatch=showLookup&contextID=" & ctxtId)
'        Call myWait(IE)
'    End If
'Next i
'
'IE.Quit
'Set IE = Nothing
'Sleep 1000
'Call checkAutocoding
'End Sub


'utility function for developers
Sub pageInfo()
' go to a page and report forms, links, anchors, etc.
Dim IE As InternetExplorerMedium ' This object (the "medium" variety as opposed to "InternetExplorer") is necessary in our security climate
Dim i As Long, n As Long
Dim x As String, currentText As String, searchText As String, oldURL As String
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
#End If 'pc
