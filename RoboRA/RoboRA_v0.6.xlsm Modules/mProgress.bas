Attribute VB_Name = "mProgress"
Sub UpdateProgressBar(PctDone As Single)
    With ufProgress
        .FrameProgress.Caption = "Completed " & VBA.Format$(PctDone, "0%")
        .LabelProgress.Width = PctDone * (.FrameProgress.Width - 10)
    End With
    DoEvents
    AppActivate Application.Caption
    DoEvents
End Sub
