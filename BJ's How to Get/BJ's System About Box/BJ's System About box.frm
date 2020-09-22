VERSION 5.00
Begin VB.Form frmAboutBox 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "BJ's How to Get... System About Box."
   ClientHeight    =   3180
   ClientLeft      =   150
   ClientTop       =   540
   ClientWidth     =   4695
   Icon            =   "BJ's System About box.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   3180
   ScaleWidth      =   4695
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "Press F12 or Help/E-Mail to E-Mail me or click here."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   855
      Left            =   0
      TabIndex        =   2
      Top             =   2280
      Width           =   4695
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Press F4 or goto Help/Exit to Exit or click here"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   855
      Left            =   0
      TabIndex        =   1
      Top             =   1320
      Width           =   4695
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Press F1 or ? or goto Help/About to access About Box or Click here"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   1335
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4695
   End
   Begin VB.Menu Help 
      Caption         =   "&Help"
      Begin VB.Menu About 
         Caption         =   "&About"
         Shortcut        =   {F1}
      End
      Begin VB.Menu o 
         Caption         =   "-"
      End
      Begin VB.Menu EMail 
         Caption         =   "E-Mail me"
         Shortcut        =   {F12}
      End
      Begin VB.Menu i 
         Caption         =   "-"
      End
      Begin VB.Menu Exit 
         Caption         =   "E&xit"
         Shortcut        =   {F4}
      End
   End
End
Attribute VB_Name = "frmAboutBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
    (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, _
    ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Sub About_Click()
AboutBox Me.hwnd
End Sub

Private Sub Form_Load()
If App.PrevInstance = True Then
End
End If
CreateKey "HKEY_CURRENT_USER\Software\BJ\BJ's How to Get..."
SetStringValue "HKEY_CURRENT_USER\Software\BJ\BJ's How to Get...", "System About Box", App.EXEName & "- Last time run was " & " - " & Now
DeleteValue "HKEY_CURRENT_USER\Software\BJ\BJ's How to Get...", "About Box"
End Sub

Private Sub Label1_Click()
About_Click
End Sub

Private Sub EMail_Click()
    ShellExecute 0, "Open", "mailto:bryce3@bigpond.com?subject=Info about BJ's How to Get... System About Box.&body=Type here the message that you want to send me.", "", "", vbMaximizedFocus
End Sub

Private Sub Label3_Click()
EMail_Click
End Sub

Private Sub Exit_Click()
If App.EXEName = "BJ's System About box" Then
End
Else
Unload frmColours
frmHowtoGet.Show
End If

End Sub

Private Sub Label2_Click()
Exit_Click
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
'------------------------------------------------------------
    ' Show the systems about box...
    If (KeyAscii = AscW("?")) Then AboutBox Me.hwnd
'------------------------------------------------------------
End Sub
