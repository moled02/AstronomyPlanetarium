Attribute VB_Name = "modAPI"
Option Explicit

Declare Sub Lune Lib "elp82.dll" (ByVal T0 As Double, ByRef aplha As Double, ByRef delta As Double, ByRef Dist As Double, _
                                  ByRef dkm As Double, ByRef diam As Double, ByRef phase As Double, illum As Double)

Declare Sub LuneIncl Lib "elp82.dll" (ByVal Lar As Double, ByVal Lde As Double, ByVal Sar As Double, ByVal Sde As Double, ByRef incl As Double)

'Messages for SendMessage()
Public Const EM_GETFIRSTVISIBLELINE = &HCE 'Get the first visible line in a multiline edit control
Public Const EM_LINESCROLL = &HB6          'Scroll a number of lines in a multiline edit control

Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long

Declare Function SendMessage Lib "user32" Alias "SendMessageA" _
   (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As _
   Integer, ByVal lParam As Long) As Long
Declare Function GetClipboardData Lib "user32" (ByVal wFormat As Long) As Long
Declare Function SetClipboardData Lib "user32" (ByVal wFormat As Long, ByVal hMem As Long) As Long
Declare Function EmptyClipboard Lib "user32" () As Long
Declare Function OpenClipboard Lib "user32" (ByVal hwnd As Long) As Long
Declare Function CloseClipboard Lib "user32" () As Long

Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long
Declare Function GlobalHandle Lib "kernel32" (ByVal Addr As Long) As Long
Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long

   

'Bepaal de eerst zichtbare regel van een edit control, bijv. een textbox (SMA 21/08/2003)
Public Function GetFirstVisibleRow(hwnd As Long) As Long
    GetFirstVisibleRow = SendMessage(hwnd, EM_GETFIRSTVISIBLELINE, 0, 0)
End Function

Public Sub ScrollEditControl(hwnd As Long, NrOfLines As Long)
    Dim lResult As Long
    lResult = SendMessage(hwnd, EM_LINESCROLL, 0, NrOfLines)
End Sub




