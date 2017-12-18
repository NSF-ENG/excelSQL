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

Private Sub UserForm_Initialize()
    ' Set the width of the progress bar to 0.
    gCancelProgress = False
    ufProgress.LabelProgress.Width = 0
'Start Userform Centered near top of Excel Screen
  Me.StartUpPosition = 0
  Me.Left = Application.Left + (Application.Width - Me.Width) / 2
  Me.Top = Application.Top + Me.Height
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
