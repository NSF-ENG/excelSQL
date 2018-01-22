VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmClearProp 
   Caption         =   "Reset/Clear RoboRA"
   ClientHeight    =   2850
   ClientLeft      =   88
   ClientTop       =   424
   ClientWidth     =   4296
   OleObjectBlob   =   "frmClearProp.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmClearProp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cbPropCancel_Click()
    Unload Me
End Sub

Private Sub cboxAll_Click()
With cboxAll
    cboxClearPropParams.Value = .Value
    cboxClearPropAddOmit.Value = .Value
    cboxClearPropData.Value = .Value
    cboxClearSavedPwd.Value = .Value
End With
End Sub

Private Sub propClear_Click()
If cboxClearPropParams.Value Then Advanced.Range("query_params").Cells.Value = HiddenSettings.Range("query_params").Cells.Value
If cboxClearPropAddOmit.Value Then Call ClearMatchingTables("*PropTable", RoboRA)
If cboxClearPropData.Value Then Call ClearQTables
If cboxClearSavedPwd.Value Then
  Prefs.Range("test_table_permissions") = ""
  Prefs.Range("dirSharedRAtemplate") = "" '\\collaboration.inside.nsf.gov@SSL\DavWWWRoot\eng\meritreview\SiteAssets\ENG Tools Websites and Best Practices\RoboRA\RAtemplates\"
  Prefs.Range("dirRAtemplate") = ""
  Prefs.Range("dirRAoutput") = ""
  RoboRA.Range("dirRAoutput2") = ""
  Advanced.Range("dirRAoutput3") = ""
  Prefs.Range("dirRoboRA") = ""
  HiddenSettings.Range("rpt_pwd").Value = ""
End If
Unload Me
End Sub

