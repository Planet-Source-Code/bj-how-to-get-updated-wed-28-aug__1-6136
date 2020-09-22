Attribute VB_Name = "modBasicCalendar"
Option Explicit

'get screen size
Dim screenx As Long
Dim screeny As Long


Public Const GCL_HCURSOR = -12

Declare Function ClipCursor Lib "user32" _
(lpRect As Any) As Long

Declare Function DestroyCursor Lib "user32" _
(ByVal hCursor As Any) As Long

Declare Function LoadCursorFromFile Lib "user32" _
Alias "LoadCursorFromFileA" (ByVal lpFileName As String) _
As Long

Declare Function SetClassLong Lib "user32" _
Alias "SetClassLongA" (ByVal hwnd As Long, _
ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Declare Function GetClassLong Lib "user32" _
Alias "GetClassLongA" (ByVal hwnd As Long, _
ByVal nIndex As Long) As Long



'taskbar stuff
Public hWnd1 As Long
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hwndInsertAfter As Long, ByVal X As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
    Public Const SWP_HIDEWINDOW = &H80
    Public Const SWP_SHOWWINDOW = &H40
    
'CHANGE RES STUFF
Private Declare Function EnumDisplaySettings Lib "user32" Alias "EnumDisplaySettingsA" (ByVal lpszDeviceName As Long, ByVal iModeNum As Long, lpDevMode As Any) As Boolean


Private Declare Function ChangeDisplaySettings Lib "user32" Alias "ChangeDisplaySettingsA" (lpDevMode As Any, ByVal dwflags As Long) As Long
    Const CCDEVICENAME = 32
    Const CCFORMNAME = 32
    Const DM_PELSWIDTH = &H80000
    Const DM_PELSHEIGHT = &H100000


Private Type DEVMODE
    dmDeviceName As String * CCDEVICENAME
    dmSpecVersion As Integer
    dmDriverVersion As Integer
    dmSize As Integer
    dmDriverExtra As Integer
    dmFields As Long
    dmOrientation As Integer
    dmPaperSize As Integer
    dmPaperLength As Integer
    dmPaperWidth As Integer
    dmScale As Integer
    dmCopies As Integer
    dmDefaultSource As Integer
    dmPrintQuality As Integer
    dmColor As Integer
    dmDuplex As Integer
    dmYResolution As Integer
    dmTTOption As Integer
    dmCollate As Integer
    dmFormName As String * CCFORMNAME
    dmUnusedPadding As Integer
    dmBitsPerPel As Integer
    dmPelsWidth As Long
    dmPelsHeight As Long
    dmDisplayFlags As Long
    dmDisplayFrequency As Long
    End Type
    Dim DevM As DEVMODE
        
'***************************************************************
' Name: Change Screen Resolution (on the fly even!)
' Description:This code allows you to change screen resolutions i
'     n win95.
' By: VB Qaid
'
'
' Inputs:

'
' Returns:None
'
'Assumes:'Example: Call ChangeRes(800,600) to change to 800 x 600
'     resolution
'
'Side Effects:None
'
'Code provided by Planet Source Code(tm) (http://www.Planet-Sourc
'     e-Code.com) 'as is', without warranties as to performance, fitnes
'     s, merchantability,and any other warranty (whether expressed or i
'     mplied).
'This source code is copyrighted by Planet Source Code who has ex
'     clusive rights to distribute it.
'It is freely redistributable for personal use in source code for
'     m, or for personal or business use in a non-source code binary ex
'     ecutable.
'All other redistributions are prohibited without express written
'     consent from Exhedra Solutions, Inc.
'***************************************************************

'Attribute VB_Name = "MODchangeRes"


Sub ChangeRes(iWidth As Single, iHeight As Single)


    Dim a As Boolean
    Dim i&
    i = 0


    Do
        a = EnumDisplaySettings(0&, i&, DevM)
        i = i + 1
    Loop Until (a = False)

    Dim b&
    DevM.dmFields = DM_PELSWIDTH Or DM_PELSHEIGHT
    DevM.dmPelsWidth = iWidth
    DevM.dmPelsHeight = iHeight
    b = ChangeDisplaySettings(DevM, 0)
End Sub

Public Sub ControlPanels(FileName As String)

Dim rtn As Double

On Error Resume Next

rtn = Shell(FileName, 4)

End Sub

Sub Main()
Load frmBasicCalender
End Sub

Public Sub CleanUp()
If App.EXEName = "BJ's Basic Calender" Then
SetStringValue "HKEY_CURRENT_USER\Software\BJ\BJ's How to Get...\BJ's Basic Calender", "Calender Text Colour", frmBasicCalender.lblMonth.ForeColor
SetStringValue "HKEY_CURRENT_USER\Software\BJ\BJ's How to Get...\BJ's Basic Calender", "Calender Back Colour", frmBasicCalender.BackColor
SetBinaryValue "HKEY_CURRENT_USER\Software\BJ\BJ's How to Get...\BJ's Basic Calender", "Calender Left", frmBasicCalender.Left
SetBinaryValue "HKEY_CURRENT_USER\Software\BJ\BJ's How to Get...\BJ's Basic Calender", "Calender Top", frmBasicCalender.Top
SetBinaryValue "HKEY_CURRENT_USER\Software\BJ\BJ's How to Get...\BJ's Basic Calender", "Calender Height", frmBasicCalender.Height
SetStringValue "HKEY_CURRENT_USER\Software\BJ\BJ's How to Get...\BJ's Basic Calender", "Show\Hide TimeZone", frmBasicCalender.mnuShowHideTimeZone.Caption
SetBinaryValue "HKEY_CURRENT_USER\Software\BJ\BJ's How to Get...\BJ's Basic Calender", "Windows Time Top", frmBasicCalender.Frame4.Top

''If screenx <> (Screen.Width / Screen.TwipsPerPixelX) Then Call ChangeRes((screenx), (screeny))
    Call Shell_NotifyIcon(NIM_DELETE, TrayIcon)
Unload frmBasicCalender
Set frmBasicCalender = Nothing
    End
    Else

'This part is used when your run my How to Get... app
SetStringValue "HKEY_CURRENT_USER\Software\BJ\BJ's How to Get...\BJ's Basic Calender", "Calender Text Colour", frmBasicCalender.lblMonth.ForeColor
SetStringValue "HKEY_CURRENT_USER\Software\BJ\BJ's How to Get...\BJ's Basic Calender", "Calender Back Colour", frmBasicCalender.BackColor
SetBinaryValue "HKEY_CURRENT_USER\Software\BJ\BJ's How to Get...\BJ's Basic Calender", "Calender Left", frmBasicCalender.Left
SetBinaryValue "HKEY_CURRENT_USER\Software\BJ\BJ's How to Get...\BJ's Basic Calender", "Calender Top", frmBasicCalender.Top
SetBinaryValue "HKEY_CURRENT_USER\Software\BJ\BJ's How to Get...\BJ's Basic Calender", "Calender Height", frmBasicCalender.Height
SetStringValue "HKEY_CURRENT_USER\Software\BJ\BJ's How to Get...\BJ's Basic Calender", "Show\Hide TimeZone", frmBasicCalender.mnuShowHideTimeZone.Caption
SetBinaryValue "HKEY_CURRENT_USER\Software\BJ\BJ's How to Get...\BJ's Basic Calender", "Windows Time Top", frmBasicCalender.Frame4.Top

''If screenx <> (Screen.Width / Screen.TwipsPerPixelX) Then Call ChangeRes((screenx), (screeny))
    Call Shell_NotifyIcon(NIM_DELETE, TrayIcon)
    'frmHowtoGet.Show
    'Unload frmBasicCalender
Unload frmBasicCalender
Set frmBasicCalender = Nothing
    End If
End Sub

