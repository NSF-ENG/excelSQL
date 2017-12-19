VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmClearSplit 
   Caption         =   "Clear Split Tracking"
   ClientHeight    =   2848
   ClientLeft      =   88
   ClientTop       =   424
   ClientWidth     =   4296
   OleObjectBlob   =   "frmClearSplit.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmClearSplit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cbSplitCancel_Click()
    Unload Me
End Sub

Private Sub splitClear_Click()
    If cboxClearSplitParams.Value Then
        ActiveSheet.Range("SplitPropParams").Cells.Value = HiddenSettings.Range("SplitPropParams").Cells.Value
        ActiveSheet.Range("SplitBudgetParams").Cells.Value = HiddenSettings.Range("SplitBudgetParams").Cells.Value
    End If
    If cboxClearSplitAddOmit.Value Then Call ClearMatchingTables("splits_*")
    If cboxClearSplitData.Value Then Call ClearMatchingTables("SplitQueryTable*")
    Unload Me
End Sub
