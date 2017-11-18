VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "HelpTab"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Function Descriptor() As clsSheetDescriptor
' return descriptor for this sheet
    Set Descriptor = New clsSheetDescriptor
    With Descriptor
     .name = Me.name
     .description = "Help documentation and use cases"
     .tip = "For more help"
     .tabtitle = "Help Documentation"
     .note = .description
     .helpText = "Including tab and button descriptions, tips, top 10 use cases, upcoming features, link to FAQ, and link to other tools."
     .helpvideo = ""
     .phase = 0 ' never refreshed
     .order = 8888
     .uigroup = .order Mod 10
     .depends = ""
    End With
End Function
