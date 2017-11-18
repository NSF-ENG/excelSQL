VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsSheetDescriptor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' Every sheet has a Descriptor function that returns its documentation and interface hints

'used in sheetTable interface (InputTab)
Public name As String ' Me.name ' taken from sheet name
Public description As String ' short description in sheet table and top of window
Public tip As String ' tooltip in SheetTable

'used in sheet heading
Public tabtitle As String ' tab title in sheet table and top of window -- will be upcased
Public note As String ' tab brief intro and help for top of window

'used in help table (HelpTab)
Public helpText As String ' text for HelpTable
Public helpvideo As String ' link to video tutorial

'disappearing soon
Public usage As String ' longer help text

'hidden in sheetTable
Public phase As Byte ' 1,2, or 3 : 1 for dependencies, 3 for long queries
Public order As Integer ' user interface order & group hint

'reserved for future use
Public uigroup As Integer ' user interface group for coloring; last digit of order at present
Public depends As String 'name of sheet depended on; currently unused. eg. InputTab.name



  
  
  
