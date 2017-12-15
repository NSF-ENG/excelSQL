Attribute VB_Name = "mProgress"
Sub UpdateProgressBar(PctDone As Single)
    With ufProgress
        .FrameProgress.Caption = "Completed " & Format(PctDone, "0%")
        .LabelProgress.Width = PctDone * (.FrameProgress.Width - 10)
    End With
    DoEvents
End Sub
