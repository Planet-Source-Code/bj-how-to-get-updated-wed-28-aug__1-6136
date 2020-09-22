VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{FC07EBD4-FE92-11D0-A199-A0077383D901}#5.5#0"; "Hackprog.ocx"
Begin VB.Form frmTrialVersion 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "BJ's How to Get... Trial Version."
   ClientHeight    =   3855
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3780
   Icon            =   "BJ's Trial Version.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   3855
   ScaleWidth      =   3780
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   195
      Left            =   45
      TabIndex        =   30
      Top             =   4440
      Width           =   3645
      _ExtentX        =   6429
      _ExtentY        =   344
      _Version        =   393216
      Appearance      =   1
      Max             =   30
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000007&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   3480
      Picture         =   "BJ's Trial Version.frx":0442
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   29
      Top             =   120
      Width           =   240
   End
   Begin CCRProgressBar.ccrpProgressBar ProgressBar2 
      Height          =   225
      Left            =   50
      Top             =   3610
      Width           =   3675
      _ExtentX        =   6482
      _ExtentY        =   397
      AutoCaption     =   2
      BackColor       =   -2147483641
      BorderStyle     =   1
      Caption         =   "30 of 30 "
      FillColor       =   255
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   65280
      IncrementSize   =   1
      Max             =   30
      ReverseFill     =   -1  'True
      Value           =   30
   End
   Begin VB.CommandButton cmdRegedit 
      Caption         =   "&Regedit"
      Height          =   495
      Left            =   840
      TabIndex        =   28
      Top             =   4800
      Width           =   1215
   End
   Begin VB.Timer Timer2 
      Interval        =   1000
      Left            =   120
      Top             =   5160
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   3240
      Top             =   5280
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      Caption         =   "                Expire Information"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   2775
      Left            =   0
      TabIndex        =   0
      Top             =   1080
      Width           =   3750
      Begin VB.Line Line4 
         BorderColor     =   &H80000009&
         X1              =   15
         X2              =   1080
         Y1              =   105
         Y2              =   105
      End
      Begin VB.Line Line2 
         BorderColor     =   &H80000009&
         X1              =   120
         X2              =   3520
         Y1              =   480
         Y2              =   480
      End
      Begin VB.Line Line1 
         BorderColor     =   &H8000000C&
         BorderWidth     =   2
         X1              =   120
         X2              =   3520
         Y1              =   480
         Y2              =   480
      End
      Begin VB.Label Label3 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Seconds"
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   2960
         TabIndex        =   17
         Top             =   240
         Width           =   630
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "60"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   2520
         TabIndex        =   16
         Top             =   240
         Width           =   375
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "BJ's Trial Version will shutdown in:"
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   240
         Width           =   2895
      End
      Begin VB.Label Label5 
         BackColor       =   &H80000002&
         BackStyle       =   0  'Transparent
         Caption         =   "Times"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   255
         Left            =   3120
         TabIndex        =   13
         Top             =   600
         Width           =   495
      End
      Begin VB.Label lblStart 
         Alignment       =   1  'Right Justify
         BackColor       =   &H0000FFFF&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Left            =   1200
         TabIndex        =   12
         Top             =   1440
         Width           =   2415
      End
      Begin VB.Label lblA 
         BackColor       =   &H80000002&
         BackStyle       =   0  'Transparent
         Caption         =   "App Started at:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Left            =   90
         TabIndex        =   11
         Top             =   1440
         Width           =   1335
      End
      Begin VB.Label lblTimes 
         Alignment       =   2  'Center
         BackColor       =   &H80000002&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF00&
         Height          =   255
         Left            =   840
         TabIndex        =   10
         Top             =   600
         Width           =   2775
      End
      Begin VB.Label lblB 
         BackColor       =   &H80000002&
         BackStyle       =   0  'Transparent
         Caption         =   "App Used:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   255
         Left            =   90
         TabIndex        =   9
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label lblC 
         BackColor       =   &H80000002&
         BackStyle       =   0  'Transparent
         Caption         =   "Begin Trial:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF00FF&
         Height          =   255
         Left            =   90
         TabIndex        =   8
         Top             =   1800
         Width           =   1215
      End
      Begin VB.Label lblD 
         BackColor       =   &H80000002&
         BackStyle       =   0  'Transparent
         Caption         =   "Trial Expiry:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   90
         TabIndex        =   7
         Top             =   2160
         Width           =   1335
      End
      Begin VB.Label lblF 
         BackColor       =   &H80000002&
         BackStyle       =   0  'Transparent
         Caption         =   "Days:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   2880
         Width           =   975
      End
      Begin VB.Label lblTrial 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FF00FF&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF00FF&
         Height          =   195
         Left            =   1200
         TabIndex        =   5
         Top             =   1800
         Width           =   2415
      End
      Begin VB.Label lblExpired 
         Alignment       =   1  'Right Justify
         BackColor       =   &H000000FF&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   1200
         TabIndex        =   4
         Top             =   2160
         Width           =   2415
      End
      Begin VB.Label lblLeft 
         BackColor       =   &H80000002&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   195
         Left            =   1200
         TabIndex        =   3
         Top             =   2880
         Width           =   2355
      End
      Begin VB.Label lblE 
         BackColor       =   &H80000002&
         BackStyle       =   0  'Transparent
         Caption         =   "Today:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   90
         TabIndex        =   2
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label lblToday 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000002&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   840
         TabIndex        =   1
         Top             =   1080
         Width           =   2775
      End
      Begin VB.Line Line3 
         BorderColor     =   &H8000000C&
         BorderWidth     =   2
         X1              =   15
         X2              =   1080
         Y1              =   105
         Y2              =   105
      End
   End
   Begin VB.Label Label11 
      Caption         =   "Label11"
      Height          =   255
      Left            =   1440
      TabIndex        =   26
      Top             =   5880
      Width           =   255
   End
   Begin VB.Label Label10 
      Caption         =   "Label10"
      Height          =   255
      Left            =   1080
      TabIndex        =   25
      Top             =   5880
      Width           =   255
   End
   Begin VB.Label Label9 
      Caption         =   "Label9"
      Height          =   255
      Left            =   720
      TabIndex        =   24
      Top             =   5880
      Width           =   255
   End
   Begin VB.Label Label8 
      Caption         =   "Label8"
      Height          =   255
      Left            =   360
      TabIndex        =   23
      Top             =   5880
      Width           =   255
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Welcome"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   240
      Left            =   0
      TabIndex        =   21
      Top             =   520
      Width           =   3735
   End
   Begin VB.Line Line6 
      BorderColor     =   &H80000009&
      X1              =   120
      X2              =   3600
      Y1              =   480
      Y2              =   480
   End
   Begin VB.Line Line5 
      BorderColor     =   &H8000000C&
      BorderWidth     =   2
      X1              =   120
      X2              =   3600
      Y1              =   480
      Y2              =   480
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "RegisteredOwner"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   255
      Left            =   0
      TabIndex        =   22
      Top             =   840
      Width           =   3735
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H80000002&
      BackStyle       =   0  'Transparent
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   2760
      TabIndex        =   19
      Top             =   120
      Width           =   735
   End
   Begin VB.Label Bryce 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "BJ's Trial Version"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   330
      Index           =   0
      Left            =   0
      TabIndex        =   14
      Top             =   170
      Width           =   2505
   End
   Begin VB.Label Bryce 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "BJ's Trial Version"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   330
      Index           =   1
      Left            =   50
      TabIndex        =   18
      Top             =   90
      Width           =   2505
   End
   Begin VB.Label Bryce 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "BJ's Trial Version"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   330
      Index           =   2
      Left            =   100
      TabIndex        =   20
      Top             =   10
      Width           =   2505
   End
   Begin VB.Label Bryce 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "BJ's Trial Version"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   330
      Index           =   3
      Left            =   150
      TabIndex        =   27
      Top             =   -70
      Width           =   2505
   End
End
Attribute VB_Name = "frmTrialVersion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
    (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, _
    ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private DateToday As Date
Private Declare Function ShellAbout Lib "shell32.dll" Alias "ShellAboutA" _
(ByVal hwnd As Long, ByVal szApp As String, ByVal szOtherStuff As String, _
ByVal hIcon As Long) As Long



Public Sub AboutBox(hwnd As Long)
'-----------------------------------------------------------------
    ' Show help about dialog...
    ShellAbout hwnd, App.EXEName, _
               vbCrLf & "BJ. E-Mail me: bryce3@bigpond.com", frmTrialVersion.Icon
'-----------------------------------------------------------------
End Sub

Private Sub Bryce_Click(Index As Integer)
Select Case Index
Case 0 To 4
ShellExecute 0, "Open", "mailto:bryce3@bigpond.com?subject=Info about BJ's Howto Get... Trial Version.&body=Type here the message that you want to send me.", "", "", vbMaximizedFocus
End Select
End Sub

Private Sub cmdRegedit_Click()
Shell "C:\Windows\Regedit.exe", vbMaximizedFocus
End Sub

Private Sub Form_Load()
If App.PrevInstance = True Then
End
End If

CreateKey "HKEY_CURRENT_USER\Software\BJ\BJ's How to Get..."
SetStringValue "HKEY_CURRENT_USER\Software\BJ\BJ's How to Get...", "Trial Version", App.EXEName & "- Last time run was " & " - " & Now

On Error GoTo ErrorHandler

Label7.Caption = GetStringValue("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion", "RegisteredOwner")

If Label7.Caption = "Error" Then
Label7.Caption = GetStringValue("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion", "RegisteredOwner")
End If

If GetStringValue("HKEY_CURRENT_USER\Software\BJ\BJ's How to Get...\Trial Version", "Times Opened") = "Error" Then
CreateKey "HKEY_CURRENT_USER\Software\BJ\BJ's How to Get...\Trial Version"
SetStringValue "HKEY_CURRENT_USER\Software\BJ\BJ's How to Get...\Trial Version", "Times Opened", "1"
End If
If GetStringValue("HKEY_CURRENT_USER\Software\BJ\BJ's How to Get...\Trial Version", "Times Opened") = "" Then
SetStringValue "HKEY_CURRENT_USER\Software\BJ\BJ's How to Get...\Trial Version", "Times Opened", "1"
End If
lblTimes.Caption = GetStringValue("HKEY_CURRENT_USER\Software\BJ\BJ's How to Get...\Trial Version", "Times Opened") & "  times"
If GetStringValue("HKEY_CURRENT_USER\Software\BJ\BJ's How to Get...\Trial Version", "Expire Date") = "Error" Then
CreateKey "HKEY_CURRENT_USER\Software\BJ\BJ's How to Get...\Trial Version"
SetStringValue "HKEY_CURRENT_USER\Software\BJ\BJ's How to Get...\Trial Version", "Expire Date", Now + 31
SetStringValue "HKEY_CURRENT_USER\Software\BJ\BJ's How to Get...\Trial Version", "Trial Start Date", Now
End If
If GetStringValue("HKEY_CURRENT_USER\Software\BJ\BJ's How to Get...\Trial Version", "Expire Date") = "" Then
SetStringValue "HKEY_CURRENT_USER\Software\BJ\BJ's How to Get...\Trial Version", "Expire Date", Now + 31
SetStringValue "HKEY_CURRENT_USER\Software\BJ\BJ's How to Get...\Trial Version", "Trial Start Date", Now
End If
lblTrial.Caption = GetStringValue("HKEY_CURRENT_USER\Software\BJ\BJ's How to Get...\Trial Version", "Trial Start Date")
lblTimes.Caption = GetStringValue("HKEY_CURRENT_USER\Software\BJ\BJ's How to Get...\Trial Version", "Times Opened")
CreateKey "HKEY_CURRENT_USER\Software\BJ\BJ's How to Get...\Trial Version"
SetStringValue "HKEY_CURRENT_USER\Software\BJ\BJ's How to Get...\Trial Version", "Times Opened", lblTimes.Caption + 1
SetStringValue "HKEY_CURRENT_USER\Software\BJ\BJ's How to Get...\Trial Version", "Todays Date", Now
lblStart.Caption = GetStringValue("HKEY_CURRENT_USER\Software\BJ\BJ's How to Get...\Trial Version", "Todays Date")
SetStringValue "HKEY_CURRENT_USER\Software\BJ\BJ's How to Get...\Trial Version", "Todays Date", lblStart.Caption
lblExpired.Caption = GetStringValue("HKEY_CURRENT_USER\Software\BJ\BJ's How to Get...\Trial Version", "Expire Date")
lblLeft = Format$(days360(lblStart.Caption, lblExpired.Caption), "###,###") & " Days"
SetStringValue "HKEY_CURRENT_USER\Software\BJ\BJ's How to Get...\Trial Version", "Days", lblLeft.Caption
lblLeft.Caption = GetStringValue("HKEY_CURRENT_USER\Software\BJ\BJ's How to Get...\Trial Version", "Days")

CreateKey "HKEY_CURRENT_USER\Software\BJ\BJ's How to Get...\Trial Version"
CreateKey "HKEY_CURRENT_USER\Software\BJ\BJ's How to Get...\Trial Version\File Info."
SetStringValue "HKEY_CURRENT_USER\Software\BJ\BJ's How to Get...\Trial Version", "", "bryce3@bigpond.com"
SetStringValue "HKEY_CURRENT_USER\Software\BJ\BJ's How to Get...\Trial Version\File Info.", "", "bryce3@bigpond.com"
SetStringValue "HKEY_CURRENT_USER\Software\BJ\BJ's How to Get...\Trial Version\File Info.", "Copyright", App.LegalCopyright
SetStringValue "HKEY_CURRENT_USER\Software\BJ\BJ's How to Get...\Trial Version\File Info.", "Trade Mark", App.LegalTrademarks
SetStringValue "HKEY_CURRENT_USER\Software\BJ\BJ's How to Get...\Trial Version\File Info.", "Version", App.Major & "." & App.Minor & "." & App.Revision & "." & "BJ"
SetStringValue "HKEY_CURRENT_USER\Software\BJ\BJ's How to Get...\Trial Version\File Info.", "File Name", App.EXEName & ".exe"
SetStringValue "HKEY_CURRENT_USER\Software\BJ\BJ's How to Get...\Trial Version\File Info.", "File Path", App.Path
SetStringValue "HKEY_CURRENT_USER\Software\BJ\BJ's How to Get...\Trial Version\File Info.", "File Description", App.FileDescription

If lblLeft = 1 & " Days" Then
lblLeft = 1 & " Day"
End If

If Val(lblLeft.Caption) < 0 Then
MsgBox "Your trial version is expired!" & vbCrLf & _
"Please get a newer version to continue." & vbCrLf & vbCrLf & _
"Good Bye...", vbOKOnly + vbCritical, "This Trial Version has Expired"
    If App.EXEName = "BJ's Trial Version" Then
    End
    Else
    Unload frmTrialVersion
    frmHowtoGet.Show
    End If
ElseIf Val(lblLeft.Caption) > 31 Then
MsgBox "Do not adjust Date/Time." & vbCrLf & _
"Your trial version is expired!" & vbCrLf & _
"Please get a newer version to continue." & vbCrLf & vbCrLf & _
"Good Bye...", vbOKOnly + vbCritical, "This Trial Version has Expired"
    If App.EXEName = "BJ's Trial Version" Then
    End
    Else
    Unload frmTrialVersion
    frmHowtoGet.Show
    End If
ElseIf Val(lblTimes.Caption) > 10 Then
MsgBox "This has now been opened more than 10 times." & vbCrLf & _
"Please get a newer version to continue." & vbCrLf & vbCrLf & _
"Good Bye...", vbOKOnly + vbCritical, "This Trial Version has Expired"
    If App.EXEName = "BJ's Trial Version" Then
    End
    Else
    Unload frmTrialVersion
    frmHowtoGet.Show
    End If
Else
If lblLeft = "1 Days" Then
lblLeft = "1 Day"
End If
If lblLeft = "Days" Then
lblLeft = "Last Day"
ProgressBar2.ForeColor = &HFF&
End If
ProgressBar1.Value = Val(lblLeft)
ProgressBar2.Value = Val(lblLeft)
ProgressBar2.Caption = lblLeft & " left from 30 Days"
End If
Exit Sub
ErrorHandler:
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label6.ForeColor = &HFFFF&
End Sub

Private Sub Form_Unload(Cancel As Integer)
SetStringValue "HKEY_CURRENT_USER\Software\BJ\BJ's How to Get...\Trial Version", "Todays Date", lblStart.Caption
End Sub

Private Sub Label6_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label6.ForeColor = &HFF&
Label2.Caption = 1
End Sub

Private Sub Picture1_Click()
AboutBox Me.hwnd
End Sub

Private Sub Timer1_Timer()
DateToday = Now
lblToday.Caption = Format$(DateToday, "dd/mm/yyyy H:mm:ss ampm")
End Sub
Public Function days360(dt1 As Date, dt2 As Date) As Long
    
    Dim z1 As Long, z2 As Long
    Dim d1 As Long, d2 As Long
    Dim m1 As Long, m2 As Long
    Dim y1 As Long, y2 As Long
    
    d1 = Day(dt1)
    m1 = Month(dt1)
    y1 = Year(dt1)
    
    d2 = Day(dt2)
    m2 = Month(dt2)
    y2 = Year(dt2)
    
    If d1 = 31 Then
        z1 = 30
    Else
        z1 = d1
    End If
    
    If d2 = 31 And d1 >= 30 Then
        z2 = 30
    Else
        z2 = d2
    End If

    days360 = (y2 - y1) * 360 + (m2 - m1) * 30 + (z2 - z1)

End Function


Private Sub Timer2_Timer()
Label2.Caption = Label2.Caption - 1
If Label2.Caption < 2 Then Label3.Caption = "Second"
If Label2.Caption = 0 Then

Label8.Caption = GetStringValue("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion", "RegisteredOwner")
If Label8.Caption = "Error" Then
Label8.Caption = GetStringValue("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion", "RegisteredOwner")
End If

Label9.Caption = GetStringValue("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion", "RegisteredOrganization")
If Label9.Caption = "Error" Then
Label9.Caption = GetStringValue("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion", "RegisteredOrganization")
End If

Label10.Caption = GetStringValue("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion", "SystemRoot")
If Label10.Caption = "Error" Then
Label10.Caption = GetStringValue("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion", "SystemRoot")
End If

Label11.Caption = GetStringValue("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion", "ProductName")
If Label11.Caption = "Error" Then
Label11.Caption = GetStringValue("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion", "ProductName")
End If

MsgBox "Thank you for trying this Trial Version." & vbCrLf & vbCrLf & _
"This has been configured and works on:" & vbCrLf & _
"Windows 95, 98, 2000, ME, XP and maybe NT" & vbCrLf & vbCrLf & _
"I hope this is what you have been looking for." & vbCrLf & _
"This Example demonstrated how to access the registry." & vbCrLf & _
"-------------------------------------------------------------" & vbCrLf & _
"Eg... Hello to:" & vbCrLf & _
"Registered Owner: " & Label8.Caption & vbCrLf & _
"Registered Organization:" & Label9.Caption & vbCrLf & _
"Your Windows Directory is: " & Label10.Caption & vbCrLf & _
"You are currently running: " & Label11.Caption & vbCrLf & _
"-------------------------------------------------------------" & vbCrLf & _
"E-Mail me if you have any problems. Thanks. BJ" & vbCrLf & _
"(bryce3@bigpond.com)", vbInformation + vbOKOnly, frmTrialVersion.Caption & " Information."
If 1 Then
    If App.EXEName = "BJ's Trial Version" Then
    End
    Else
    Unload frmTrialVersion
    frmHowtoGet.Show
    End If
End If
End If
End Sub
