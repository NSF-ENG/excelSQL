VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Favorites"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Function Descriptor() As clsSheetDescriptor
' return descriptor for this sheet
    Set Descriptor = New clsSheetDescriptor
    With Descriptor
     .name = Me.name
     .description = "Saves parameter and tab refresh settings"
     .tip = "Can edit " & .name
     .tabtitle = "Favorite Settings"
     .note = .description
     .helpText = "Holds favorite parameters and tab refresh setting; you may edit them directly if you are careful."
     .helpvideo = ""
     .phase = 0 ' never refreshed
     .order = 9989 ' near last
     .uigroup = .order Mod 10
     .depends = ""
    End With
End Function
