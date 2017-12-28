VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ufProgress 
   Caption         =   "RoboRA Progress Bar"
   ClientHeight    =   688
   ClientLeft      =   96
   ClientTop       =   416
   ClientWidth     =   4288
   OleObjectBlob   =   "ufProgress.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ufProgress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' to use: ufProgress.Show vbModeless
'Private Sub UpdateProgressBar(PctDone As Single)
'    ufProgress
'        .FrameProgress.Caption = "RoboRA Progress " & Format(PctDone, "0%")
'        .LabelProgress.Width = PctDone * (.FrameProgress.Width - 10)
'    End With
'    DoEvents
'End Sub
' when done: unload ufProgress

'Note: this userform does not always stay on top of excel when it is "not responding" because a query is long.
' For PC, the following will probably help
' http://www.jkp-ads.com/Articles/keepuserformontop02.asp
' https://www.mrexcel.com/forum/excel-questions/386643-userform-always-top.html
' http://www.vbaexpress.com/forum/showthread.php?5071-Solved-Keep-userform-on-top-at-all-times
' For Mac, i need to check if this is a problem before trying to solve it.

Private Sub UserForm_Initialize()
    ' Set the width of the progress bar to 0.
    gCancelProgress = False
    ufProgress.LabelProgress.Width = 0
    ' Start Userform Centered in Excel Screen
    Me.StartUpPosition = 0
    Me.Left = Application.Left + (Application.Width - Me.Width) / 2
    Me.Top = Application.Top + (Application.Height - Me.Height) / 2
End Sub


Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
  'Intercept/repurpose Unload if user clicks form "X" close button.
  If CloseMode = 0 Then
    If MsgBox("Do you want to abort the current sequence of actions?  May leave partial results." & vbNewLine _
              & "The current action may need to complete before aborting sequence.", vbYesNo) <> vbYes Then
      Cancel = True
    Else
      gCancelProgress = True
    End If
  End If
End Sub
