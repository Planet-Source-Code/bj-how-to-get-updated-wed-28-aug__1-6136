VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmMakeINI 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "BJ's How to Get... Make INI"
   ClientHeight    =   8115
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6720
   Icon            =   "BJ's Make INI.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   8115
   ScaleWidth      =   6720
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkDelete 
      Caption         =   "Delete .INI file"
      Height          =   495
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   240
      Width           =   1215
   End
   Begin VB.CommandButton cmdAbout 
      Caption         =   "About"
      Height          =   495
      Left            =   3600
      TabIndex        =   3
      Top             =   240
      Width           =   1215
   End
   Begin VB.CommandButton cmdEMail 
      Caption         =   "E-Mail me"
      Height          =   495
      Left            =   1920
      TabIndex        =   2
      Top             =   240
      Width           =   1215
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   495
      Left            =   5280
      TabIndex        =   1
      Top             =   240
      Width           =   1215
   End
   Begin RichTextLib.RichTextBox rtbShowINI 
      Height          =   7095
      Left            =   120
      TabIndex        =   0
      Top             =   960
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   12515
      _Version        =   393217
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   3
      DisableNoScroll =   -1  'True
      TextRTF         =   $"BJ's Make INI.frx":0442
   End
End
Attribute VB_Name = "frmMakeINI"
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
               vbCrLf & "BJ. E-Mail me: bryce3@bigpond.com", frmMakeINI.Icon
'-----------------------------------------------------------------
End Sub


Private Sub chkDelete_Click()
If chkDelete.Value = 1 Then
chkDelete.Caption = "Make .INI file"
Kill "C:\Windows\" & INI_FILE
rtbShowINI.Text = ""
End If
If chkDelete.Value = 0 Then
chkDelete.Caption = "Delete .INI file"

Call Form_Load
End If
End Sub

Private Sub cmdAbout_Click()
AboutBox Me.hwnd
End Sub

Private Sub cmdEMail_Click()
    ShellExecute 0, "Open", "mailto:bryce3@bigpond.com?subject=Info about BJ's How to Get... Make INI.&body=Type here the message that you want to send me.", "", "", vbMaximizedFocus
End Sub

Private Sub cmdExit_Click()
If App.EXEName = "BJ's Make INI" Then
If chkDelete.Caption = "Make .INI file" Then
End
Else
Kill "C:\Windows\" & INI_FILE
End
End If
Else
If chkDelete.Caption = "Make .INI file" Then
Unload frmMakeINI
frmHowtoGet.Show
Else
Kill "C:\Windows\" & INI_FILE
Unload frmMakeINI
frmHowtoGet.Show
End If
End If
End Sub

Private Sub Form_DblClick()
Shell "C:\Windows\Notepad.exe C:\Windows\BJ'sHo~1.ini", vbNormalFocus
End Sub

Private Sub rtbShowINI_DblClick()
Shell "C:\Windows\Notepad.exe C:\Windows\BJ'sHo~1.ini", vbNormalFocus
End Sub

Private Sub Form_Load()
If App.PrevInstance = True Then
End
End If
CreateKey "HKEY_CURRENT_USER\Software\BJ\BJ's How to Get..."
SetStringValue "HKEY_CURRENT_USER\Software\BJ\BJ's How to Get...", "Make .INI", App.EXEName & "- Last time run was " & " - " & Now
    chkDelete.Value = 0
    BJ = WritePrivateProfileString(SECTION, Entry, " bryce3@bigpond.com", INI_FILE)
        BJ = GetPrivateProfileString(SECTION, Entry, " bryce3@bigpond.com", BJ, Len(BJ), INI_FILE)
    
    BJ = WritePrivateProfileString(SECTION, ENTRY1, " Entry1", INI_FILE)
        BJ = GetPrivateProfileString(SECTION, ENTRY1, " Entry1", BJ, Len(BJ), INI_FILE)
    
    BJ = WritePrivateProfileString(SECTION, ENTRY2, " Entry2" & vbNewLine, INI_FILE)
        BJ = GetPrivateProfileString(SECTION, ENTRY2, " Entry2", BJ, Len(BJ), INI_FILE)
    
    BJ = WritePrivateProfileString(SECTION1, ENTRY3, " Double click here or form to open the .ini file in your Windows Directory" & vbNewLine, INI_FILE)
        BJ = GetPrivateProfileString(SECTION1, ENTRY3, " Entry3", BJ, Len(BJ), INI_FILE)
    
    BJ = WritePrivateProfileString(SECTION2, ENTRY4, " " & "C:\Windows\" & INI_FILE, INI_FILE)
        BJ = GetPrivateProfileString(SECTION2, ENTRY4, " " & App.Path, BJ, Len(BJ), INI_FILE)
    
    BJ = WritePrivateProfileString(SECTION2, ENTRY5, " " & App.Path, INI_FILE)
        BJ = GetPrivateProfileString(SECTION2, ENTRY5, " " & App.Path, BJ, Len(BJ), INI_FILE)
    
    BJ = WritePrivateProfileString(SECTION2, ENTRY6, " " & App.EXEName, INI_FILE)
        BJ = GetPrivateProfileString(SECTION2, ENTRY6, " " & App.EXEName, BJ, Len(BJ), INI_FILE)
    
    BJ = WritePrivateProfileString(SECTION2, ENTRY7, " " & App.Major & "." & App.Minor & "." & App.Revision & vbNewLine, INI_FILE)
        BJ = GetPrivateProfileString(SECTION2, ENTRY7, " " & App.Major & "." & App.Minor & "." & App.Revision, BJ, Len(BJ), INI_FILE)
    
    BJ = WritePrivateProfileString(SECTION3, ENTRY8, " " & 1, INI_FILE)
        BJ = GetPrivateProfileString(SECTION3, ENTRY8, " ", BJ, Len(BJ), INI_FILE)
    
    BJ = WritePrivateProfileString(SECTION3, ENTRY9, " " & 2, INI_FILE)
        BJ = GetPrivateProfileString(SECTION3, ENTRY9, " ", BJ, Len(BJ), INI_FILE)
    
    BJ = WritePrivateProfileString(SECTION3, ENTRY10, " " & 3, INI_FILE)
        BJ = GetPrivateProfileString(SECTION3, ENTRY10, " ", BJ, Len(BJ), INI_FILE)
    
    BJ = WritePrivateProfileString(SECTION3, ENTRY11, " " & 4, INI_FILE)
        BJ = GetPrivateProfileString(SECTION3, ENTRY11, " ", BJ, Len(BJ), INI_FILE)
    
    BJ = WritePrivateProfileString(SECTION3, ENTRY12, " " & 5, INI_FILE)
        BJ = GetPrivateProfileString(SECTION3, ENTRY12, " ", BJ, Len(BJ), INI_FILE)
    
    BJ = WritePrivateProfileString(SECTION3, ENTRY13, " " & 6, INI_FILE)
        BJ = GetPrivateProfileString(SECTION3, ENTRY13, " ", BJ, Len(BJ), INI_FILE)
    
    BJ = WritePrivateProfileString(SECTION3, ENTRY14, " " & 7, INI_FILE)
        BJ = GetPrivateProfileString(SECTION3, ENTRY14, " ", BJ, Len(BJ), INI_FILE)
    
    BJ = WritePrivateProfileString(SECTION3, ENTRY15, " " & 8, INI_FILE)
        BJ = GetPrivateProfileString(SECTION3, ENTRY15, " ", BJ, Len(BJ), INI_FILE)
    
    BJ = WritePrivateProfileString(SECTION3, ENTRY16, " " & 9, INI_FILE)
        BJ = GetPrivateProfileString(SECTION3, ENTRY16, " ", BJ, Len(BJ), INI_FILE)
    
    BJ = WritePrivateProfileString(SECTION3, ENTRY17, " " & 10, INI_FILE)
        BJ = GetPrivateProfileString(SECTION3, ENTRY17, " ", BJ, Len(BJ), INI_FILE)
    
    BJ = WritePrivateProfileString(SECTION3, ENTRY18, " " & 11, INI_FILE)
        BJ = GetPrivateProfileString(SECTION3, ENTRY18, " ", BJ, Len(BJ), INI_FILE)
    
    BJ = WritePrivateProfileString(SECTION3, ENTRY19, " " & 12, INI_FILE)
        BJ = GetPrivateProfileString(SECTION3, ENTRY19, " ", BJ, Len(BJ), INI_FILE)
    
    BJ = WritePrivateProfileString(SECTION3, ENTRY20, " " & 13, INI_FILE)
        BJ = GetPrivateProfileString(SECTION3, ENTRY20, " ", BJ, Len(BJ), INI_FILE)
    
    BJ = WritePrivateProfileString(SECTION3, ENTRY21, " " & 14, INI_FILE)
        BJ = GetPrivateProfileString(SECTION3, ENTRY21, " ", BJ, Len(BJ), INI_FILE)
    
    BJ = WritePrivateProfileString(SECTION3, ENTRY22, " " & 15, INI_FILE)
        BJ = GetPrivateProfileString(SECTION3, ENTRY22, " ", BJ, Len(BJ), INI_FILE)
    
    BJ = WritePrivateProfileString(SECTION3, ENTRY23, " " & 16, INI_FILE)
        BJ = GetPrivateProfileString(SECTION3, ENTRY23, " ", BJ, Len(BJ), INI_FILE)
        
    BJ = WritePrivateProfileString(SECTION3, ENTRY24, " " & 17, INI_FILE)
        BJ = GetPrivateProfileString(SECTION3, ENTRY24, " ", BJ, Len(BJ), INI_FILE)
    
rtbShowINI.FileName = "C:\Windows\" & INI_FILE

End Sub

Private Sub cmdShow_Click()
On Error Resume Next
rtbShowINI.FileName = "C:\Windows\" & INI_FILE
If rtbShowINI.Text = "Double Click to Delete .ini file" Then
If rtbShowINI.Text = "" Then
If rtbShowINI.Text = " " Then
MsgBox "Please click on Make INI first.", vbInformation, "No INI file."
Else
rtbShowINI.FileName = "C:\Windows\" & INI_FILE
End If
End If
End If
End Sub
