Attribute VB_Name = "mCleanCopy"
Option Explicit
#If Mac Then
' use this on mac; use windows api on PC.  May change this so both use api calls
'Private
Sub CopyText(Text As String)
    'VBA Macro using late binding to copy text to clipboard.
    'By Justin Kay, 8/15/2014
    'Thanks to http://akihitoyamashiro.com/en/VBA/LateBindingDataObject.htm
    'Needs reference MS Office Object Library
    Dim cb As Object
    
    #If Mac Then
        Set cb = New DataObject
    #Else
        Set cb = CreateObject("new:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")
    #End If
    cb.Clear
    cb.SetText Text
    cb.PutInClipboard
    Set cb = Nothing
End Sub

#Else 'PC
' Clipboard via windows API. https://www.thespreadsheetguru.com/blog/2015/1/13/how-to-use-vba-code-to-copy-text-to-the-clipboard
' Handle 64-bit and 32-bit Office:
#If VBA7 Then
  Declare PtrSafe Function GlobalUnlock Lib "kernel32" (ByVal hMem As LongPtr) As Long
  Declare PtrSafe Function GlobalLock Lib "kernel32" (ByVal hMem As LongPtr) As Long
  Declare PtrSafe Function GlobalAlloc Lib "kernel32" (ByVal wFlags As LongPtr, ByVal dwBytes As LongPtr) As Long
  Declare PtrSafe Function CloseClipboard Lib "User32" () As Long
  Declare PtrSafe Function OpenClipboard Lib "User32" (ByVal hwnd As LongPtr) As Long
  Declare PtrSafe Function EmptyClipboard Lib "User32" () As Long
  Declare PtrSafe Function lstrcpy Lib "kernel32" (ByVal lpString1 As Any, ByVal lpString2 As Any) As Long
  Declare PtrSafe Function SetClipboardData Lib "User32" (ByVal wFormat As LongPtr, ByVal hMem As LongPtr) As Long
#Else
  Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
  Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
  Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
  Declare Function CloseClipboard Lib "User32" () As Long
  Declare Function OpenClipboard Lib "User32" (ByVal hwnd As Long) As Long
  Declare Function EmptyClipboard Lib "User32" () As Long
  Declare Function lstrcpy Lib "kernel32" (ByVal lpString1 As Any, ByVal lpString2 As Any) As Long
  Declare Function SetClipboardData Lib "User32" (ByVal wFormat As Long, ByVal hMem As Long) As Long
#End If

Public Const GHND = &H42
Public Const CF_TEXT = 1
Public Const MAXSIZE = 4096

Sub CopyText(MyString As String)
'PURPOSE: API function to copy text to clipboard
'SOURCE: www.msdn.microsoft.com/en-us/library/office/ff192913.aspx

Dim hGlobalMemory As Long, lpGlobalMemory As Long
Dim hClipMemory As Long, x As Long
  hGlobalMemory = GlobalAlloc(GHND, Len(MyString) + 1) 'Allocate moveable global memory
  lpGlobalMemory = GlobalLock(hGlobalMemory) 'Lock the block to get a far pointer to this memory.
  lpGlobalMemory = lstrcpy(lpGlobalMemory, MyString) 'Copy the string to this global memory.
  If GlobalUnlock(hGlobalMemory) <> 0 Then 'Unlock the memory.
    MsgBox "Could not unlock memory location. Copy aborted."
    GoTo OutOfHere2
  End If
  If OpenClipboard(0&) = 0 Then 'Open the Clipboard to copy data to.
    MsgBox "Could not open the Clipboard. Copy aborted."
    Exit Sub
  End If
  x = EmptyClipboard() 'Clear the Clipboard.
  hClipMemory = SetClipboardData(CF_TEXT, hGlobalMemory) 'Copy the data to the Clipboard.
OutOfHere2:
  If CloseClipboard() = 0 Then MsgBox "Could not close Clipboard."
End Sub
#End If

Function FixIPSText(s As String) As String
'
' Replace special characters
'

Dim fromC As String, toC As String
Dim i As Long
'Do not edit & save on Mac or it will munge the special characters
'I need to fix this.
fromC = VBA.ChrW$(160) & "–”“’‘•ãèéàáâåçêëìíîïòóôõùúû"
toC = " -""""''*aeeaaaaceeiiiioooouuu"

'MsgBox Len(fromC) & " : " & Len(toC)
    s = Replace(Replace(Replace(Replace(Replace(s, "…", "..."), "—", "--"), "ä", "ae"), "ö", "oe"), "ü", "ue") ' mulitcharacter replaces —äöü

    For i = 1 To Len(fromC)
       'MsgBox vba.AscW(vba.Mid(fromC, i, 1)) & " - " & vba.AscW(vba.Mid(toC, i, 1))
       s = Replace(s, VBA.Mid$(fromC, i, 1), VBA.Mid$(toC, i, 1))
    Next i
    For i = 1 To Len(s) ' replace anything I missed with ?
      If VBA.AscW(VBA.Mid$(s, i, 1)) > 127 Then s = Replace(s, VBA.Mid$(s, i, 1), "?")
    Next i
    s = Replace(s, VBA.Chr$(11), VBA.Chr$(13)) ' fix vertical tab to CR
    s = Replace(Replace(s, "{{", ""), "}}", "") ' strip double braces, which is text included in RAs but omitted in POcomments
FixIPSText = s
End Function

Function StripDoubleBrackets(s As String) As String
' does not handle nesting.
Dim i As Long, j As Long, k As Long, lenS As Long

j = InStrRev(s, "]]")
Do While j > 0
    i = InStrRev(s, "[[", j)
    k = InStrRev(s, "]]", j - 1)
    If i < k Then
      MsgBox "Warning: Consecutive close comment brackets with no open." & vbNewLine & VBA.Mid$(s, k, j - k + 1)
      j = k
    End If
    If i < 1 Then
      MsgBox "Warning: missing open comment brackets for first close." & vbNewLine & VBA.Left$(s, j)
    Else
      s = VBA.Left$(s, i - 1) & VBA.Right$(s, Len(s) - j - 1)
    End If
    j = InStrRev(s, "]]", i)
Loop

StripDoubleBrackets = s
End Function


