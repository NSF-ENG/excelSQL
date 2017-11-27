VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} PropClearFrm 
   Caption         =   "Clear Proposal Tracking"
   ClientHeight    =   2848
   ClientLeft      =   96
   ClientTop       =   416
   ClientWidth     =   4288
   OleObjectBlob   =   "PropClearFrm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "PropClearFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cbPropCancel_Click()
    Debug.Print (ActiveSheet.Name)
    Unload Me
End Sub

Private Sub propClear_Click()
Debug.Print (ActiveSheet.Name)
If cboxClearPropParams.Value Then ActiveSheet.Range("PropParams").Cells.Value = HiddenSettings.Range("PropParams").Cells.Value
If cboxClearPropAddOmit.Value Then Call ClearMatchingTable("props_*")
If cboxClearPropData.Value Then Call ClearMatchingTable("PropQueryTable*")
End Sub
