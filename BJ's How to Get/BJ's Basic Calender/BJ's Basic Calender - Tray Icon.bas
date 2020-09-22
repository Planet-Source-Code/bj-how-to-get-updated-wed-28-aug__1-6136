Attribute VB_Name = "modTrayIcon"
Option Explicit



'**Originally published by  Ryan Heldt (rheldt@vb-online.com)
'**Modified by Donovan Parks (donopark@awinc.com)

'Win32 API declaration
Declare Function Shell_NotifyIcon Lib "Shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean

' Constants used to detect clicking on the icon
Public Const WM_LBUTTONDBLCLK = &H203
Public Const WM_RBUTTONUP = &H205

' Constants used to control the icon
Public Const NIM_ADD = &H0
Public Const NIM_MODIFY = &H1
Public Const NIF_MESSAGE = &H1
Public Const NIM_DELETE = &H2
Public Const NIF_ICON = &H2
Public Const NIF_TIP = &H4

' Used as the ID of the call back message
Public Const WM_MOUSEMOVE = &H200

' Used by Shell_NotifyIcon
Public Type NOTIFYICONDATA
    cbSize As Long
    hwnd As Long
    uId As Long
    uFlags As Long
    ucallbackMessage As Long
    hIcon As Long
    szTip As String * 64
End Type

'create variable of type NOTIFYICONDATA
Public TrayIcon As NOTIFYICONDATA

Public Declare Function ShellAbout Lib "shell32.dll" Alias "ShellAboutA" _
(ByVal hwnd As Long, ByVal szApp As String, ByVal szOtherStuff As String, _
ByVal hIcon As Long) As Long
