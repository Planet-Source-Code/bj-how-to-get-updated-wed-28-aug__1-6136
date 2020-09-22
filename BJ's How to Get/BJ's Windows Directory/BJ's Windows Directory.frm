VERSION 5.00
Begin VB.Form frmWindowsDir 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "BJ's How to Get... Windows Directory."
   ClientHeight    =   1530
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4035
   Icon            =   "BJ's Windows Directory.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   1530
   ScaleWidth      =   4035
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdAbout 
      Caption         =   "About"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1440
      TabIndex        =   3
      Top             =   1080
      Width           =   1215
   End
   Begin VB.CommandButton cmdEMail 
      Caption         =   "E-Mail me"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   1080
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2760
      TabIndex        =   0
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H80000012&
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   600
      Width           =   3855
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H80000012&
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   3855
   End
End
Attribute VB_Name = "frmWindowsDir"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Copy all this text you are reading. Select all
' by clicking on the text
' hold mouse button down and go to the end of the Text
' Now press CTRL + C to Copy
' Create new EXE project and paste it into Form1
' Now You will need
' 1: 2 labels named Label1 and Label2 which are the "Default Names"
' 2: 3 Command Buttons named Command1. which is the "Default Name"
' Change Caption to Get Windows Directory
' 3: Command Button named Command2. which is the "Default Name"
' Change Caption to Exit
' 3: Command Button named Command2. which is the "Default Name"
' Change Caption to About
' Now click Run or F5 and Click on Command Button

'Function to get Windows directory
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
    (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, _
    ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Declare Function GetWindowsDirectory Lib "kernel32" Alias _
"GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long

Private Declare Function GetSystemDirectory Lib "kernel32" Alias _
"GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long


'Variable to store the Windows directory
Dim WinDir As String
Dim WinSysDir As String

'Buffer and constant used for API functions
Dim msBuffer As String * 255
Const BUFFERSIZE = 255
Private Declare Function ShellAbout Lib "shell32.dll" Alias "ShellAboutA" _
(ByVal hwnd As Long, ByVal szApp As String, ByVal szOtherStuff As String, _
ByVal hIcon As Long) As Long



Public Sub AboutBox(hwnd As Long)
'-----------------------------------------------------------------
    ' Show help about dialog...
    ShellAbout hwnd, App.EXEName, _
               vbCrLf & "BJ. E-Mail me: bryce3@bigpond.com", frmWindowsDir.Icon
'-----------------------------------------------------------------
End Sub

Private Sub cmdAbout_Click()
AboutBox Me.hwnd
End Sub

Private Sub cmdEMail_Click()
    ShellExecute 0, "Open", "mailto:bryce3@bigpond.com?subject=Info about BJ's How to Get... Windows Directory.&body=Type here the message that you want to send me.", "", "", vbMaximizedFocus
End Sub

Private Sub Command2_Click() 'Exit Command Button
If App.EXEName = "BJ's Windows Directory" Then
End
Else
Unload frmWindowsDir
frmHowtoGet.Show
End If
End Sub


Private Sub Form_Load()
If App.PrevInstance = True Then
End
End If
CreateKey "HKEY_CURRENT_USER\Software\BJ\BJ's How to Get..."
SetStringValue "HKEY_CURRENT_USER\Software\BJ\BJ's How to Get...", "Windows Directory", App.EXEName & "- Last time run was " & " - " & Now
Dim lBytes As Long
Dim lBytes1 As Long

lBytes = GetWindowsDirectory(msBuffer, BUFFERSIZE)
lBytes1 = GetSystemDirectory(msBuffer, BUFFERSIZE)
WinDir = Left$(msBuffer, lBytes)
WinSysDir = Left$(msBuffer, lBytes1)

Label1.Caption = WinDir  'Which = above Default = C:\Windows
Label2.Caption = WinSysDir  'Which = above Default = C:\Windows\System

End Sub
