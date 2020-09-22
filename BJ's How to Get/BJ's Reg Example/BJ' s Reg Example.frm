VERSION 5.00
Begin VB.Form frmRegExample 
   Caption         =   "BJ's How to Get... Reg Example"
   ClientHeight    =   5640
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6975
   Icon            =   "BJ' s Reg Example.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5640
   ScaleWidth      =   6975
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdRegEdit 
      Caption         =   "1 - 1 - Regedit"
      Height          =   495
      Left            =   3600
      TabIndex        =   15
      Top             =   120
      Width           =   3255
   End
   Begin VB.CommandButton cmdEMail 
      Caption         =   "7 - E-Mail"
      Height          =   495
      Left            =   3600
      TabIndex        =   14
      Top             =   4440
      Width           =   3255
   End
   Begin VB.CommandButton cmdAbout 
      Caption         =   "6 - About"
      Height          =   495
      Left            =   120
      TabIndex        =   13
      Top             =   4440
      Width           =   3255
   End
   Begin VB.CommandButton cmdDeleteDWordValue 
      Caption         =   "4 - 2 - Delete DWord Value"
      Height          =   495
      Left            =   4680
      TabIndex        =   12
      Top             =   1920
      Width           =   2175
   End
   Begin VB.CommandButton cmdDeleteBinaryValue 
      Caption         =   "3 - 2 - Delete Binary Value"
      Height          =   495
      Left            =   2400
      TabIndex        =   11
      Top             =   1920
      Width           =   2175
   End
   Begin VB.CommandButton cmdDeleteStringValue 
      Caption         =   "2 - 2 - Delete String Value"
      Height          =   495
      Left            =   120
      TabIndex        =   10
      Top             =   1920
      Width           =   2175
   End
   Begin VB.TextBox Text1 
      Height          =   1215
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   9
      Top             =   3120
      Width           =   6735
   End
   Begin VB.CommandButton cmdGetDWordValue 
      Caption         =   "4 - 1 - Get DWord Value"
      Height          =   495
      Left            =   4680
      TabIndex        =   8
      Top             =   1320
      Width           =   2175
   End
   Begin VB.CommandButton cmdGetBinaryValue 
      Caption         =   "3 - 1 - Get Binary Value"
      Height          =   495
      Left            =   2400
      TabIndex        =   7
      Top             =   1320
      Width           =   2175
   End
   Begin VB.CommandButton cmdGetStringValue 
      Caption         =   "2 -1 - Get String Value"
      Height          =   495
      Left            =   120
      TabIndex        =   6
      Top             =   1320
      Width           =   2175
   End
   Begin VB.CommandButton cmdCreateDWordValue 
      Caption         =   "4 - Create DWord Value"
      Height          =   495
      Left            =   4680
      TabIndex        =   5
      Top             =   720
      Width           =   2175
   End
   Begin VB.CommandButton cmdCreateBinaryValue 
      Caption         =   "3 - Create Binary Value"
      Height          =   495
      Left            =   2400
      TabIndex        =   4
      Top             =   720
      Width           =   2175
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "8 - Exit"
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Top             =   5040
      Width           =   6735
   End
   Begin VB.CommandButton cmdDeleteKey 
      Caption         =   "5 - Delete Key"
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   2520
      Width           =   6735
   End
   Begin VB.CommandButton cmdCreateStringValue 
      Caption         =   "2 - Create String Value"
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   2175
   End
   Begin VB.CommandButton cmdCreateKey 
      Caption         =   "1 - Create Key"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3255
   End
End
Attribute VB_Name = "frmRegExample"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function ShellAbout Lib "shell32.dll" Alias "ShellAboutA" _
(ByVal hwnd As Long, ByVal szApp As String, ByVal szOtherStuff As String, _
ByVal hIcon As Long) As Long

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
    (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, _
    ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long



Private Sub AboutBox(hwnd As Long)
'-----------------------------------------------------------------
    ' Show help about dialog...
    ShellAbout hwnd, App.EXEName, _
               vbCrLf & "BJ. E-Mail me: bryce3@bigpond.com", frmRegExample.Icon
'-----------------------------------------------------------------
End Sub

Private Sub cmdCreateKey_Click() 'Create Key
CreateKey "HKEY_CURRENT_USER\Software\BJ\BJ's How to Get...\BJ's Registry Example"
Text1.Text = "HKEY_CURRENT_USER\Software\BJ\BJ's How to Get...\BJ's Registry Example" & vbNewLine & "was created for you."

End Sub

Private Sub cmdAbout_Click() 'About
AboutBox Me.hwnd

End Sub

Private Sub cmdEMail_Click() 'E-Mail
    ShellExecute 0, "Open", "mailto:bryce3@bigpond.com?subject=Info about BJ's How to Get... Reg Example.&body=Type here the message that you want to send me.", "", "", vbMaximizedFocus

End Sub

Private Sub cmdRegEdit_Click() 'Registry Editor
Shell "C:\Windows\regedit.exe", vbMaximizedFocus
End Sub

Private Sub cmdCreateStringValue_Click() 'Create String Value
SetStringValue "HKEY_CURRENT_USER\Software\BJ\BJ's How to Get...\BJ's Registry Example", "String", "This is a String Test by BJ. I'm just testing if this works or not."
Text1.Text = "The String Value 'String' with the Value 'This is a String Test by BJ. I'm just testing if this works or not.' was created for you"

End Sub

Private Sub cmdGetStringValue_Click() 'Get String Value
Text1.Text = GetStringValue("HKEY_CURRENT_USER\Software\BJ\BJ's How to Get...\BJ's Registry Example", "String")

End Sub

Private Sub cmdDeleteStringValue_Click() 'Delete String Value
DeleteValue "HKEY_CURRENT_USER\Software\BJ\BJ's How to Get...\BJ's Registry Example", "String"
Text1.Text = "If you click on any of the Create buttons you will get an Error because the value has been deleted"

End Sub

Private Sub cmdCreateBinaryValue_Click() 'Create Binary Value
SetBinaryValue "HKEY_CURRENT_USER\Software\BJ\BJ's How to Get...\BJ's Registry Example", "Binary", "This is a Binary Test by BJ. I'm just testing if this works or not."
Text1.Text = "The Binary Value 'Binary' with the Value 'This is a Binary Test by BJ. I'm just testing if this works or not.' was created for you"

End Sub

Private Sub cmdGetBinaryValue_Click() 'Get Binary Value
Text1.Text = GetBinaryValue("HKEY_CURRENT_USER\Software\BJ\BJ's How to Get...\BJ's Registry Example", "Binary")
End Sub

Private Sub cmdDeleteBinaryValue_Click() 'Delete Binary Value
DeleteValue "HKEY_CURRENT_USER\Software\BJ\BJ's How to Get...\BJ's Registry Example", "Binary"
Text1.Text = "If you click on any of the Create buttons you will get an Error because the value has been deleted"

End Sub

Private Sub cmdCreateDWordValue_Click() 'Create DWord Value
SetDWORDValue "HKEY_CURRENT_USER\Software\BJ\BJ's How to Get...\BJ's Registry Example", "DWord", "20"
Text1.Text = "The DWord Value 'DWord' with the Value '' was created for you" & _
vbNewLine & _
vbNewLine & _
"If anyone knows how to write DWords propley, can you E-Mail me."

End Sub

Private Sub cmdGetDWordValue_Click() 'Get DWord Value
Text1.Text = GetBinaryValue("HKEY_CURRENT_USER\Software\BJ\BJ's How to Get...\BJ's Registry Example", "DWord")

End Sub

Private Sub cmdDeleteDWordValue_Click() 'Delete DWord Value
DeleteValue "HKEY_CURRENT_USER\Software\BJ\BJ's How to Get...\BJ's Registry Example", "DWord"
Text1.Text = "If you click on any of the Create buttons you will get an Error because the value has been deleted"

End Sub

Private Sub cmdDeleteKey_Click() 'Delete Key
DeleteKey "HKEY_CURRENT_USER\Software\BJ\BJ's How to Get...", "BJ's Registry Example"
Text1.Text = "If you click on any of the Create buttons you will get an Error because the Key has been deleted"

End Sub

Private Sub cmdExit_Click() 'Exit
End

End Sub

Private Sub Form_Load()
If App.PrevInstance = True Then
End
End If
CreateKey "HKEY_CURRENT_USER\Software\BJ\BJ's How to Get..."
SetStringValue "HKEY_CURRENT_USER\Software\BJ\BJ's How to Get...", "Registry Example", App.EXEName & "- Last time run was " & " - " & Now


End Sub
