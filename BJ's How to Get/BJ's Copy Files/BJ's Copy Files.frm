VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCopyFile 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "BJ's How to Get... Copy Files."
   ClientHeight    =   3135
   ClientLeft      =   1695
   ClientTop       =   1515
   ClientWidth     =   4830
   Icon            =   "BJ's Copy Files.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3135
   ScaleWidth      =   4830
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdEMail 
      Caption         =   "E-Mail me"
      Height          =   375
      Left            =   1440
      TabIndex        =   11
      Top             =   2520
      Width           =   975
   End
   Begin VB.CommandButton cmdAbout 
      Caption         =   "About"
      Height          =   375
      Left            =   360
      TabIndex        =   10
      Top             =   2520
      Width           =   975
   End
   Begin VB.Frame Frame 
      Height          =   2895
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4575
      Begin MSComctlLib.ProgressBar ProgressBar 
         Height          =   375
         Left            =   240
         TabIndex        =   12
         Top             =   1920
         Width           =   4215
         _ExtentX        =   7435
         _ExtentY        =   661
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.CommandButton Browsedestination 
         Caption         =   "Bro&wse..."
         Enabled         =   0   'False
         Height          =   375
         Left            =   3480
         TabIndex        =   6
         Top             =   1080
         Width           =   975
      End
      Begin VB.TextBox Destinationpath 
         Enabled         =   0   'False
         Height          =   285
         Left            =   120
         TabIndex        =   5
         Top             =   1080
         Width           =   3255
      End
      Begin VB.CommandButton Browsefile 
         Caption         =   "&Browse..."
         Height          =   375
         Left            =   3480
         TabIndex        =   3
         Top             =   480
         Width           =   975
      End
      Begin VB.TextBox Filepath 
         Height          =   285
         Left            =   120
         TabIndex        =   2
         Top             =   480
         Width           =   3255
      End
      Begin VB.CommandButton Cancel 
         Caption         =   "E&xit"
         Height          =   375
         Left            =   3480
         TabIndex        =   9
         Top             =   2400
         Width           =   975
      End
      Begin VB.CommandButton Copy 
         Caption         =   "Copy &File"
         Height          =   375
         Left            =   2400
         TabIndex        =   8
         Top             =   2400
         Width           =   975
      End
      Begin MSComDlg.CommonDialog Dialog 
         Left            =   3840
         Top             =   0
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         Flags           =   6148
      End
      Begin VB.Label Destinationlabel 
         Caption         =   "&Destination:"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   840
         Width           =   855
      End
      Begin VB.Label Filelabel 
         Caption         =   "&Source:"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Percentlabel 
         Caption         =   "Percent complete:"
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   1550
         Width           =   1335
      End
   End
End
Attribute VB_Name = "frmCopyFile"
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
               vbCrLf & "BJ. E-Mail me: bryce3@bigpond.com", frmCopyFile.Icon
'-----------------------------------------------------------------
End Sub

Function CopyFile(Src As String, Dst As String) As Single

Static Buf$
Dim BTest!, FSize!
Dim Chunk%, F1%, F2%

Const BUFSIZE = 1024

If Len(Dir(Dst)) Then
   Response = MsgBox(Dst + Chr(10) + Chr(10) + "File already exists. Do you want to overwrite it?", vbYesNo + vbQuestion) 'prompt the user with a message box
   If Response = vbNo Then
      Exit Function
   Else
      Kill Dst
   End If
End If
 
On Error GoTo FileCopyError
F1 = FreeFile
Open Src For Binary As F1
F2 = FreeFile
Open Dst For Binary As F2
 
FSize = LOF(F1)
BTest = FSize - LOF(F2)

Do
If BTest < BUFSIZE Then
   Chunk = BTest
Else
   Chunk = BUFSIZE
End If
      
Buf = String(Chunk, " ")
Get F1, , Buf
Put F2, , Buf
BTest = FSize - LOF(F2)

ProgressBar.Value = (100 - Int(100 * BTest / FSize))

Loop Until BTest = 0
Close F1
Close F2
CopyFile = FSize
ProgressBar.Value = 0
Exit Function

FileCopyError:
MsgBox "Copy Error!, Please try again..."
Close F1
Close F2
Exit Function

End Function

Public Function ExtractName(SpecIn As String) As String
   
Dim i As Integer
Dim SpecOut As String
   
On Error Resume Next
   
For i = Len(SpecIn) To 1 Step -1
If Mid(SpecIn, i, 1) = "\" Then
   SpecOut = Mid(SpecIn, i + 1)
   Exit For
End If
Next i

ExtractName = SpecOut

End Function



Private Sub Browsedestination_Click()

Dim bi As BROWSEINFO
Dim rtn&, pidl&, path$, pos%

bi.hOwner = Me.hwnd
bi.lpszTitle = "Browse for Destination..."
bi.ulFlags = BIF_RETURNONLYFSDIRS
pidl& = SHBrowseForFolder(bi)
  
path = Space(512)
T = SHGetPathFromIDList(ByVal pidl&, ByVal path)
pos% = InStr(path$, Chr$(0))
SpecIn = Left(path$, pos - 1)
If Right$(SpecIn, 1) = "\" Then
   SpecOut = SpecIn
Else
   SpecOut = SpecIn + "\"
End If

Destinationpath.Text = SpecOut + ExtractName(Filepath.Text)
    
End Sub

Private Sub Browsefile_Click()

Dialog.DialogTitle = "Browse for source..."
Dialog.Filter = "Any File (*.*)|*.*"
Dialog.ShowOpen
Filepath.Text = Dialog.FileName

End Sub

Private Sub Cancel_Click()
If App.EXEName = "BJ's Copy Files" Then
End
Else

Unload frmCopyFile
frmHowtoGet.Show
End If
End Sub



Private Sub cmdAbout_Click()
AboutBox Me.hwnd

End Sub

Private Sub cmdEMail_Click()
    ShellExecute 0, "Open", "mailto:bryce3@bigpond.com?subject=Info about BJ's How to Get... Copy Files.&body=Type here the message that you want to send me.", "", "", vbMaximizedFocus

End Sub

Private Sub Copy_Click()

On Error Resume Next

If Filepath.Text = "" Then
   MsgBox "You must specify a file and path in the text box provided", vbCritical
   Exit Sub
End If
If Destinationpath.Text = "" Then
   MsgBox "You must specify a destination path in the text box provided", vbCritical
   Exit Sub
End If

ProgressBar.Value = CopyFile(Filepath.Text, Destinationpath.Text)
MsgBox "You have copied " & vbNewLine & _
Filepath.Text & vbNewLine & _
" - to - " & vbNewLine & _
Destinationpath.Text
End Sub


Private Sub filepath_Change()

Destinationpath.Enabled = True
Browsedestination.Enabled = True
Destinationpath.SetFocus

End Sub


Private Sub Form_Load()
If App.PrevInstance = True Then
End
End If

CreateKey "HKEY_CURRENT_USER\Software\BJ\BJ's How to Get..."
SetStringValue "HKEY_CURRENT_USER\Software\BJ\BJ's How to Get...", "Copy Files", App.EXEName & "- Last time run was " & " - " & Now

Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 2


End Sub



