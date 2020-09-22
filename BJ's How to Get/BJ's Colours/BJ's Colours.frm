VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmColors 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "BJ's How to Get... Colours."
   ClientHeight    =   2370
   ClientLeft      =   690
   ClientTop       =   1830
   ClientWidth     =   4905
   Icon            =   "BJ's Colours.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2370
   ScaleWidth      =   4905
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   4320
      Top             =   1320
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdCustomColor 
      Caption         =   "Custom Colors"
      Height          =   495
      Left            =   2520
      TabIndex        =   21
      Top             =   1800
      Width           =   975
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   495
      Left            =   3840
      TabIndex        =   19
      Top             =   1800
      Width           =   975
   End
   Begin VB.CommandButton cmdEMail 
      Caption         =   "E-Mail"
      Height          =   495
      Left            =   1200
      TabIndex        =   18
      Top             =   1800
      Width           =   975
   End
   Begin VB.CommandButton cmdAbout 
      Caption         =   "About"
      Height          =   495
      Left            =   120
      TabIndex        =   17
      Top             =   1800
      Width           =   975
   End
   Begin VB.PictureBox picColorArr 
      Height          =   495
      Index           =   15
      Left            =   4320
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   15
      Top             =   720
      Width           =   495
   End
   Begin VB.PictureBox picColorArr 
      Height          =   495
      Index           =   14
      Left            =   3720
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   14
      Top             =   720
      Width           =   495
   End
   Begin VB.PictureBox picColorArr 
      Height          =   495
      Index           =   13
      Left            =   3120
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   13
      Top             =   720
      Width           =   495
   End
   Begin VB.PictureBox picColorArr 
      Height          =   495
      Index           =   12
      Left            =   2520
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   12
      Top             =   720
      Width           =   495
   End
   Begin VB.PictureBox picColorArr 
      Height          =   495
      Index           =   11
      Left            =   1920
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   11
      Top             =   720
      Width           =   495
   End
   Begin VB.PictureBox picColorArr 
      Height          =   495
      Index           =   10
      Left            =   1320
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   10
      Top             =   720
      Width           =   495
   End
   Begin VB.PictureBox picColorArr 
      Height          =   495
      Index           =   9
      Left            =   720
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   9
      Top             =   720
      Width           =   495
   End
   Begin VB.PictureBox picColorArr 
      Height          =   495
      Index           =   8
      Left            =   120
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   8
      Top             =   720
      Width           =   495
   End
   Begin VB.PictureBox picColorArr 
      Height          =   495
      Index           =   7
      Left            =   4320
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   7
      Top             =   120
      Width           =   495
   End
   Begin VB.PictureBox picColorArr 
      Height          =   495
      Index           =   6
      Left            =   3720
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   6
      Top             =   120
      Width           =   495
   End
   Begin VB.PictureBox picColorArr 
      Height          =   495
      Index           =   5
      Left            =   3120
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   5
      Top             =   120
      Width           =   495
   End
   Begin VB.PictureBox picColorArr 
      Height          =   495
      Index           =   4
      Left            =   2520
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   4
      Top             =   120
      Width           =   495
   End
   Begin VB.PictureBox picColorArr 
      Height          =   495
      Index           =   3
      Left            =   1920
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   3
      Top             =   120
      Width           =   495
   End
   Begin VB.PictureBox picColorArr 
      Height          =   495
      Index           =   2
      Left            =   1320
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   2
      Top             =   120
      Width           =   495
   End
   Begin VB.PictureBox picColorArr 
      Height          =   495
      Index           =   1
      Left            =   720
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   1
      Top             =   120
      Width           =   495
   End
   Begin VB.PictureBox picColorArr 
      Height          =   495
      Index           =   0
      Left            =   120
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   0
      Top             =   120
      Width           =   495
   End
   Begin VB.PictureBox Picture1 
      Height          =   495
      Left            =   2760
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   16
      Top             =   4560
      Width           =   495
   End
   Begin VB.Label lblcolorclick 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Click Color to choose color"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   20
      Top             =   1320
      Width           =   4695
   End
End
Attribute VB_Name = "frmColors"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
    (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, _
    ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Declare Function ShellAbout Lib "shell32.dll" Alias "ShellAboutA" _
(ByVal hwnd As Long, ByVal szApp As String, ByVal szOtherStuff As String, _
ByVal hIcon As Long) As Long



Public Sub AboutBox(hwnd As Long)
'-----------------------------------------------------------------
    ' Show help about dialog...
    ShellAbout hwnd, App.EXEName, _
               vbCrLf & "BJ. E-Mail me: bryce3@bigpond.com", frmColors.Icon
'-----------------------------------------------------------------
End Sub

Private Sub cmdAbout_Click()
AboutBox Me.hwnd
End Sub

Private Sub cmdCustomColor_Click()
CommonDialog1.ShowColor
    Me.BackColor = CommonDialog1.Color
End Sub

Private Sub cmdEMail_Click()
    ShellExecute 0, "Open", "mailto:bryce3@bigpond.com?subject=Info about BJ's How to Get... Colors.&body=Type here the message that you want to send me.", "", "", vbMaximizedFocus

End Sub

Private Sub cmdExit_Click()
If App.EXEName = "BJ's Colours" Then
End
Else
Unload frmColours
frmHowtoGet.Show
End If
End Sub

'*************************************************
' Purpose:  Unload the form.
'*************************************************
'*************************************************
' Purpose:  Initialize the form by setting the
'           colors of the picture boxes.
'*************************************************
Private Sub Form_Load()
If App.PrevInstance = True Then
End
End If
    Dim intI As Integer ' counter
CreateKey "HKEY_CURRENT_USER\Software\BJ\BJ's How to Get..."
SetStringValue "HKEY_CURRENT_USER\Software\BJ\BJ's How to Get...", "Colours", App.EXEName & "- Last time run was " & " - " & Now
    For intI = 0 To 15 '16 colors
        picColorArr(intI).BackColor = QBColor(intI)
    Next intI
End Sub
'*************************************************
' Purpose:  Sets the text color of the selection
'           on the calling form.
' Inputs:   intIndex: The index of the clicked pict.
'*************************************************
Private Sub picColorArr_Click(intIndex As Integer)
    
    Me.BackColor = QBColor(intIndex)
End Sub
