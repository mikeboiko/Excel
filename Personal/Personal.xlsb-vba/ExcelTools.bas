Attribute VB_Name = "ExcelTools"
' Put Personal.xlsb into C:\Users\<username>\AppData\Roaming\Microsoft\Excel\XLSTART
' Then open another workbook and click View -> Hide

'Handle 64-bit and 32-bit Office
#If VBA7 Then
  Declare PtrSafe Function GlobalUnlock Lib "kernel32" (ByVal hMem As LongPtr) As Long
  Declare PtrSafe Function GlobalLock Lib "kernel32" (ByVal hMem As LongPtr) As Long
  Declare PtrSafe Function GlobalAlloc Lib "kernel32" (ByVal wFlags As LongPtr, _
    ByVal dwBytes As LongPtr) As Long
  Declare PtrSafe Function CloseClipboard Lib "User32" () As Long
  Declare PtrSafe Function OpenClipboard Lib "User32" (ByVal hwnd As LongPtr) As Long
  Declare PtrSafe Function EmptyClipboard Lib "User32" () As Long
  Declare PtrSafe Function lstrcpy Lib "kernel32" (ByVal lpString1 As Any, _
    ByVal lpString2 As Any) As Long
  Declare PtrSafe Function SetClipboardData Lib "User32" (ByVal wFormat _
    As LongPtr, ByVal hMem As LongPtr) As Long
#Else
  Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
  Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
  Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, _
    ByVal dwBytes As Long) As Long
  Declare Function CloseClipboard Lib "User32" () As Long
  Declare Function OpenClipboard Lib "User32" (ByVal hwnd As Long) As Long
  Declare Function EmptyClipboard Lib "User32" () As Long
  Declare Function lstrcpy Lib "kernel32" (ByVal lpString1 As Any, _
    ByVal lpString2 As Any) As Long
  Declare Function SetClipboardData Lib "User32" (ByVal wFormat _
    As Long, ByVal hMem As Long) As Long
#End If

Public Const GHND = &H42
Public Const CF_TEXT = 1
Public Const MAXSIZE = 4096

Function ClipBoard_SetData(MyString As String)
'PURPOSE: API function to copy text to clipboard
'SOURCE: www.msdn.microsoft.com/en-us/library/office/ff192913.aspx

Dim hGlobalMemory As Long, lpGlobalMemory As Long
Dim hClipMemory As Long, X As Long

'Allocate moveable global memory
  hGlobalMemory = GlobalAlloc(GHND, Len(MyString) + 1)

'Lock the block to get a far pointer to this memory.
  lpGlobalMemory = GlobalLock(hGlobalMemory)

'Copy the string to this global memory.
  lpGlobalMemory = lstrcpy(lpGlobalMemory, MyString)

'Unlock the memory.
  If GlobalUnlock(hGlobalMemory) <> 0 Then
    MsgBox "Could not unlock memory location. Copy aborted."
    GoTo OutOfHere2
  End If

'Open the Clipboard to copy data to.
  If OpenClipboard(0&) = 0 Then
    MsgBox "Could not open the Clipboard. Copy aborted."
    Exit Function
  End If

'Clear the Clipboard.
  X = EmptyClipboard()

'Copy the data to the Clipboard.
  hClipMemory = SetClipboardData(CF_TEXT, hGlobalMemory)

OutOfHere2:
  If CloseClipboard() = 0 Then
    MsgBox "Could not close Clipboard."
  End If

End Function

' Copy text in selected cell to Windows clipboard
Sub CopyTextToClipboard()
Attribute CopyTextToClipboard.VB_ProcData.VB_Invoke_Func = "y\n14"
'PURPOSE: Copy a given text to the clipboard (using Windows API)
'SOURCE: www.TheSpreadsheetGuru.com
'NOTES: Must have above API declaration and ClipBoard_SetData function in your code
On Error Resume Next

Dim txt As String

'Put some text inside a string variable
  txt = Selection.Value

'Place text into the Clipboard
   ClipBoard_SetData txt

'Notify User
 ' MsgBox "There is now text copied to your clipboard!", vbInformation

End Sub


' Paste as Values into current selection
Sub PasteValues()
Attribute PasteValues.VB_ProcData.VB_Invoke_Func = "V\n14"
    On Error Resume Next
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
End Sub

' This macro removes any filtering in order to display all of the data
Sub AutoFilter_Remove()
Attribute AutoFilter_Remove.VB_ProcData.VB_Invoke_Func = "U\n14"
    On Error Resume Next
    ' Cells.AutoFilter
    ActiveSheet.ShowAllData
End Sub

