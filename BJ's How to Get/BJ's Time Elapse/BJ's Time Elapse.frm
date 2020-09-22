VERSION 5.00
Begin VB.Form frm_Time_Elapse 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "BJ's How to Get... Time Elapse"
   ClientHeight    =   2490
   ClientLeft      =   5325
   ClientTop       =   4380
   ClientWidth     =   4830
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   12
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "BJ's Time Elapse.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2490
   ScaleWidth      =   4830
   Begin VB.CommandButton cmdAbout 
      Caption         =   "About"
      Height          =   495
      Left            =   2400
      TabIndex        =   10
      Top             =   1920
      Width           =   2295
   End
   Begin VB.CommandButton cmdEMail 
      Caption         =   "E-Mail"
      Height          =   495
      Left            =   120
      TabIndex        =   9
      Top             =   1920
      Width           =   2295
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   1320
      Width           =   1575
   End
   Begin VB.CommandButton cmdEnd 
      Caption         =   "&End Timing"
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   1575
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "&StartTiming"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label lblElapsed 
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   495
      Left            =   3120
      TabIndex        =   8
      Top             =   1320
      Width           =   1575
   End
   Begin VB.Label lblEnd 
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   495
      Left            =   3120
      TabIndex        =   7
      Top             =   720
      Width           =   1575
   End
   Begin VB.Label lblStart 
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   495
      Left            =   3120
      TabIndex        =   6
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "Elapsed Time"
      ForeColor       =   &H00800000&
      Height          =   615
      Left            =   1680
      TabIndex        =   5
      Top             =   1320
      Width           =   1455
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "End Time"
      ForeColor       =   &H00008000&
      Height          =   375
      Left            =   1800
      TabIndex        =   4
      Top             =   720
      Width           =   1335
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Start Time"
      ForeColor       =   &H00000080&
      Height          =   375
      Left            =   1800
      TabIndex        =   3
      Top             =   120
      Width           =   1335
   End
End
Attribute VB_Name = "frm_Time_Elapse"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
    (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, _
    ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long


Dim StartTime As Variant
Dim EndTime As Variant
Dim ElapsedTime As Variant
Private Declare Function ShellAbout Lib "shell32.dll" Alias "ShellAboutA" _
(ByVal hwnd As Long, ByVal szApp As String, ByVal szOtherStuff As String, _
ByVal hIcon As Long) As Long



Public Sub AboutBox(hwnd As Long)
'-----------------------------------------------------------------
    ' Show help about dialog...
    ShellAbout hwnd, App.EXEName, _
               vbCrLf & "BJ. E-Mail me: bryce3@bigpond.com", frm_Time_Elapse.Icon
'-----------------------------------------------------------------
End Sub

Private Sub cmdAbout_Click()
AboutBox Me.hwnd
End Sub

Private Sub cmdEMail_Click()
ShellExecute 0, "Open", "mailto:bryce3@bigpond.com?subject=Info about BJ's How to Get... Time Elapse.&body=Type here the message that you want to send me.", "", "", vbMaximizedFocus
End Sub

Private Sub cmdEnd_Click()
'Find the ending time, compute the elapsed time
'Put both values in label boxes
EndTime = Now
ElapsedTime = EndTime - StartTime
lblEnd.Caption = Format(EndTime, "hh:mm:ss")
lblElapsed.Caption = Format(ElapsedTime, "hh:mm:ss")
End Sub

Private Sub cmdExit_Click()
End
End Sub

Private Sub cmdStart_Click()
'Establish and print starting time
StartTime = Now
lblStart.Caption = Format(StartTime, "hh:mm:ss")
lblEnd.Caption = ""
lblElapsed.Caption = ""
End Sub


Private Sub Form_Load()
CreateKey "HKEY_CURRENT_USER\Software\BJ\BJ's How to Get..."
SetStringValue "HKEY_CURRENT_USER\Software\BJ\BJ's How to Get...", "Time Elapse", App.EXEName & "- Last time run was " & " - " & Now

End Sub
