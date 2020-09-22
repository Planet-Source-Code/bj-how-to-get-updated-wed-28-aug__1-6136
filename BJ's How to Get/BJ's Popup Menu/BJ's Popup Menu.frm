VERSION 5.00
Begin VB.Form frmPopup_Menu 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "BJ's How to Get... Popup Menu."
   ClientHeight    =   585
   ClientLeft      =   -10005
   ClientTop       =   1575
   ClientWidth     =   6675
   Icon            =   "BJ's Popup Menu.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   585
   ScaleWidth      =   6675
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Right Click for Popup Menu"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   615
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6615
   End
   Begin VB.Menu mnuPopup 
      Caption         =   "Popupmenu"
      Visible         =   0   'False
      Begin VB.Menu mnuFile 
         Caption         =   "&File"
         Begin VB.Menu mnuFileNew 
            Caption         =   "&New"
            Shortcut        =   ^N
         End
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "&Edit"
         Begin VB.Menu mnuEditCopy 
            Caption         =   "&Copy"
            Shortcut        =   ^C
         End
      End
      Begin VB.Menu mnuHelp 
         Caption         =   "&Help"
         Begin VB.Menu mnuHelpAbout 
            Caption         =   "&About..."
         End
         Begin VB.Menu mnuHelpEMail 
            Caption         =   "E-&Mail"
         End
      End
      Begin VB.Menu mnuFileBar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
      End
   End
End
Attribute VB_Name = "frmPopup_Menu"
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
               vbCrLf & "BJ. E-Mail me: bryce3@bigpond.com", frmPopup_Menu.Icon
'-----------------------------------------------------------------
End Sub

Private Sub Form_DblClick()
PopupMenu mnuPopup
End Sub

Private Sub Label1_DblClick()
PopupMenu mnuPopup

End Sub

Private Sub Form_Load()
If App.PrevInstance = True Then
End
End If
CreateKey "HKEY_CURRENT_USER\Software\BJ\BJ's How to Get..."
SetStringValue "HKEY_CURRENT_USER\Software\BJ\BJ's How to Get...", "Popup Menu", App.EXEName & "- Last time run was " & " - " & Now
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button And vbRightButton _
        Then PopupMenu mnuPopup
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button And vbRightButton _
        Then PopupMenu mnuPopup
        
If Button And vbMiddleButton _
        Then PopupMenu mnuPopup
End Sub

Private Sub mnuFileExit_Click()
If App.EXEName = "BJ's Popup Menu" Then
End
Else
Unload frmPopup_Menu
frmHowtoGet.Show
End If
End Sub

Private Sub mnuFileNew_Click()
  MsgBox "New File Code goes here!"
End Sub

Private Sub mnuEditCopy_Click()
  MsgBox "Place Copy Code here!"
End Sub

Private Sub mnuHelpAbout_Click()
AboutBox Me.hwnd
End Sub

Private Sub mnuHelpEMail_Click()
    ShellExecute 0, "Open", "mailto:bryce3@bigpond.com?subject=Info about BJ's How to Get... Popup Menu.&body=Type here the message that you want to send me.", "", "", vbMaximizedFocus
End Sub
