Attribute VB_Name = "mProgress"
Option Explicit
' support progress bar
' gCancelProgress = true means to cancel the upcoming computation
' it may be set by clicking the X in the progress bar userform

Global gCancelProgress As Boolean

Sub UpdateProgressBar(PctDone As Single)
    If gCancelProgress Then End ' hard abort of current computation
    With ufProgress
        .FrameProgress.Caption = "Completed " & VBA.Format$(PctDone, "0%")
        .LabelProgress.Width = PctDone * (.FrameProgress.Width - 10)
    End With
    DoEvents
    AppActivate Application.Caption
    DoEvents
End Sub
