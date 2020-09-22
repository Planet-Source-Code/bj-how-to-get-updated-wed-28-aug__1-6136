VERSION 5.00
Begin VB.Form frmDisable_Close_Button 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "BJ's How to Get... Disable Close Button [X]."
   ClientHeight    =   1125
   ClientLeft      =   2265
   ClientTop       =   2415
   ClientWidth     =   4440
   Icon            =   "BJ's Disable Close Button.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   1125
   ScaleWidth      =   4440
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Click to Close"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   615
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   4455
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Click for About"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   720
      Width           =   1935
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Click to E-Mail me."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF00FF&
      Height          =   375
      Left            =   2040
      TabIndex        =   0
      Top             =   720
      Width           =   2295
   End
End
Attribute VB_Name = "frmDisable_Close_Button"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
    (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, _
    ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Declare Function DeleteMenu Lib "User32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Private Declare Function GetSystemMenu Lib "User32" (ByVal hwnd As Long, ByVal bRevert As Long) As Long

Private Const MF_BYPOSITION = &H400&

Private ReadyToClose As Boolean
Private Declare Function ShellAbout Lib "shell32.dll" Alias "ShellAboutA" _
(ByVal hwnd As Long, ByVal szApp As String, ByVal szOtherStuff As String, _
ByVal hIcon As Long) As Long



Public Sub AboutBox(hwnd As Long)
'-----------------------------------------------------------------
    ' Show help about dialog...
    ShellAbout hwnd, App.EXEName, _
               vbCrLf & "BJ. E-Mail me: bryce3@bigpond.com", frmDisable_Close_Button.Icon
'-----------------------------------------------------------------
End Sub
Private Sub RemoveMenus(frm As Form, remove_close As Boolean)
Dim hMenu As Long
    
    ' Get the form's system menu handle.
    hMenu = GetSystemMenu(hwnd, False)
'    If remove_close Then DeleteMenu hMenu, 6, MF_BYPOSITION ' Removes Close
'    If remove_close Then DeleteMenu hMenu, 5, MF_BYPOSITION ' Removes Seperator between Maximize and Close
'    If remove_close Then DeleteMenu hMenu, 1, MF_BYPOSITION ' Removes Move
'    If remove_close Then DeleteMenu hMenu, 4, MF_BYPOSITION ' Removes Maximize but not the Button
'    If remove_close Then DeleteMenu hMenu, 3, MF_BYPOSITION ' Removes Minimize but not the Button
'    If remove_close Then DeleteMenu hMenu, 2, MF_BYPOSITION ' Removes
'    If remove_close Then DeleteMenu hMenu, 0, MF_BYPOSITION ' Removes Remove
    
    If remove_close Then DeleteMenu hMenu, 6, MF_BYPOSITION
    End Sub

Private Sub Form_Load()
If App.PrevInstance = True Then
End
End If
    RemoveMenus Me, True
CreateKey "HKEY_CURRENT_USER\Software\BJ\BJ's How to Get..."
SetStringValue "HKEY_CURRENT_USER\Software\BJ\BJ's How to Get...", "Disable [X]", App.EXEName & "- Last time run was " & " - " & Now

End Sub

Private Sub Label1_Click()
AboutBox Me.hwnd
End Sub

Private Sub Label2_Click()
    ReadyToClose = True
If App.EXEName = "BJ's Disable Close Button" Then
End
Else
Unload frmDisable_Close_Button
frmHowtoGet.Show
End If

End Sub

Private Sub Label3_Click()
    ShellExecute 0, "Open", "mailto:bryce3@bigpond.com?subject=Info about BJ's How to Get... Disable Close Button [X].&body=Type here the message that you want to send me.", "", "", vbMaximizedFocus
End Sub
