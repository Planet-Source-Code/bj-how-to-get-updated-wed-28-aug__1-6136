VERSION 5.00
Begin VB.Form frmEMail 
   Caption         =   "BJ's How to Get... E-Mail"
   ClientHeight    =   495
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   3720
   Icon            =   "BJ's E-Mail.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   495
   ScaleWidth      =   3720
   StartUpPosition =   3  'Windows Default
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3000
      TabIndex        =   2
      Top             =   0
      Width           =   615
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "About"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1560
      TabIndex        =   1
      Top             =   0
      Width           =   1095
   End
   Begin VB.Label lblEMail 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "E-Mail"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   1095
   End
End
Attribute VB_Name = "frmEMail"
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
               vbCrLf & "BJ. E-Mail me: bryce3@bigpond.com", frmEMail.Icon
'-----------------------------------------------------------------
End Sub

Private Sub Form_Load()
If App.PrevInstance = True Then
End
End If
CreateKey "HKEY_CURRENT_USER\Software\BJ\BJ's How to Get..."
SetStringValue "HKEY_CURRENT_USER\Software\BJ\BJ's How to Get...", "E-Mail", App.EXEName & "- Last time run was " & " - " & Now


End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblEMail.ForeColor = &H80000012
End Sub

Private Sub Label1_Click()
AboutBox Me.hwnd

End Sub

Private Sub Label2_Click()
End
End Sub

Private Sub lblEMail_Click()
    ShellExecute 0, "Open", "mailto:bryce3@bigpond.com?subject=Info about BJ's How to Get... E-Mail.&body=Type here the message that you want to send me.", "", "", vbMaximizedFocus

End Sub

Private Sub lblEMail_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblEMail.ForeColor = &H80&       'Dark Blue
End Sub

Private Sub lblEMail_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblEMail.ForeColor = &H8000&    'Dark Green

End Sub
