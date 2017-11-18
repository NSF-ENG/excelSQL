Attribute VB_Name = "Module_Favorites"

Option Explicit
' handle setting and saving of params
' Favorites page holds default params in B51:C70 reserves D for tab refresh
' Last use params are E-G
' Favorites H-J, K-M, N-P, Q-S 'JSS to be revised.

Sub ddFavorites_Change()
' handle favorite list.  Params 0=Reset, 1=previous, 2-3 reserved, 4-? favorites
  'Application.ScreenUpdating = False
  Application.EnableEvents = False
  
  With InputTab.Shapes("ddFavorites").ControlFormat
    Select Case .ListIndex

      Case 1 ' Reset parameters
        Call SaveParams(2) ' save for undo
        Call SetParams(1, False)

      Case 2 ' Restore Previous Parameters
        Call SetParams(2)

      Case 3 ' Save To End of Favorites List
        Dim s As Variant
        s = InputBox("Save parameters as favorite:", "Save Favorite", "name")
        If s <> False Then 'user didn't cancel
            If Len(s) = 0 Or s = "name" Then s = "Favorite#" & .ListCount - 3 ' make up name if we didn't have one
            Dim lr As ListRow
            Set lr = Favorites.ListObjects("FavoritesMenu").ListRows.Add(AlwaysInsert:=True)
            lr.Range.Value = s ' add name of favorite as last and save favorite
            Call SaveParams(.ListCount)
        End If

      Case 4 ' Delete from Favorites
        If .ListCount < 5 Then
           MsgBox "No stored favorites available to delete"
        ElseIf MsgBox("Currently, only deletion of last is implemented", vbOKCancel) = vbOK Then
          Favorites.ListObjects("FavoritesMenu").ListRows(.ListCount).Delete 'delete last table row
        End If
        
      Case Else ' use indicated favorite
        If .ListCount < .ListIndex Then MsgBox ("Error Index > Count. Ignoring & crossing my fingers.")
        Call SaveParams(2) ' save for undo
        Call SetParams(.ListIndex)
    End Select
    .ListIndex = 0 ' Reset to blank
  End With
  Application.EnableEvents = True
 ' Application.ScreenUpdating = True
End Sub

Private Sub SaveParams(fav As Integer)
 ' save current interface parameters as favorites in columns 2*fav+1,2*fav+2
 ' input parameters at row 51, tab refresh values at 71, and add/omit tables at 101
 ' Note: this limits the interface to 20 parameters, 30 tabs
 
  Dim rng As Range
  Dim lo As ListObject
  With Favorites.Cells(11, 2 * fav + 1)   ' one-cell range at upper corner of destination
    Set rng = InputTab.Range("parameters") ' origin: assume formatting at dest is the same (text or date)
    .Resize(rng.Rows.Count, rng.Columns.Count).Cells.Value = rng.Cells.Value 'destination matches origin size and copies values;
    Set rng = InputTab.Range("SheetTable[[Ref?]:[Limit]]") 'origin Refresh flags/limits
    .Offset(30, 0).Resize(rng.Rows.Count, rng.Columns.Count).Cells.Value = rng.Cells.Value 'copy
    Set lo = InputTab.ListObjects("AddProjIDTable")
    .Offset(70, 0).Cells.Value = lo.ListRows.Count ' save # of rows
    If lo.ListRows.Count > 0 Then .Offset(71, 0).Resize(lo.ListRows.Count, 1).Cells.Value = lo.DataBodyRange.Cells.Value 'copy
    Set lo = InputTab.ListObjects("OmitPropIDTable")
    .Offset(70, 1).Cells.Value = lo.ListRows.Count
    If lo.ListRows.Count > 0 Then .Offset(71, 1).Resize(lo.ListRows.Count, 1).Cells.Value = lo.DataBodyRange.Cells.Value 'copy
  End With
End Sub

Private Sub SetParams(fav As Integer, Optional setTabs As Boolean = True)
' From favorites range fav,
'   set parameters, tab refresh (If setTabs is false, does parameters only)
' fav: 1=reset, 2=previous, 3,4 unused, 5- user favorites

  Dim rng As Range
  Dim lo As ListObject
  With Favorites.Cells(11, 2 * fav + 1) ' one-cell range at upper corner of origin
    Set rng = InputTab.Range("parameters") ' destination: assume formatting at orig is the same (text or date)
    rng.Cells.Value = .Resize(rng.Rows.Count, rng.Columns.Count).Cells.Value 'orig matches dest size and copies values
    If setTabs Then ' update tabs only if set-tabs is set
        Set rng = InputTab.Range("SheetTable[[Ref?]:[Limit]]") 'origin Refresh flags/limits
        rng.Cells.Value = .Offset(30, 0).Resize(rng.Rows.Count, rng.Columns.Count).Cells.Value 'copy
        Call hideSheets ' change which sheets are hidden
    End If
    Set lo = InputTab.ListObjects("AddProjIDTable")
    Call ClearTable(lo)
    If .Offset(70, 0).Cells.Value > 0 Then ' change prop_id table only if non-empty one saved
        lo.ListRows.Add
        .Offset(71, 0).Resize(.Offset(70, 0).Cells.Value, 1).Copy
        lo.ListRows(1).Range.PasteSpecial xlPasteValues
    End If
     Set lo = InputTab.ListObjects("OmitPropIDTable")
     Call ClearTable(lo)
     If .Offset(70, 1).Cells.Value > 0 Then ' change prop_id table only if non-empty one saved
        lo.ListRows.Add
        .Offset(71, 1).Resize(.Offset(70, 1).Cells.Value, 1).Copy
        lo.ListRows(1).Range.PasteSpecial xlPasteValues
    End If
  End With
End Sub



