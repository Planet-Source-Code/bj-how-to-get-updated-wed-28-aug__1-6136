VERSION 5.00
Begin VB.Form frmDeleteFiles 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "BJ's How to get... Delete Files."
   ClientHeight    =   3450
   ClientLeft      =   3855
   ClientTop       =   1755
   ClientWidth     =   6225
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3450
   ScaleWidth      =   6225
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Height          =   3495
      Left            =   3800
      TabIndex        =   1
      Top             =   -60
      Width           =   2415
      Begin VB.PictureBox Picture2 
         AutoSize        =   -1  'True
         Height          =   540
         Left            =   1800
         Picture         =   "BJ's Delete Files.frx":0000
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   10
         Top             =   1440
         Visible         =   0   'False
         Width           =   540
      End
      Begin VB.CommandButton cmdEMail 
         Caption         =   "E-Mail"
         Height          =   375
         Left            =   120
         TabIndex        =   9
         Top             =   2520
         Width           =   1575
      End
      Begin VB.CommandButton cmdAbout 
         Caption         =   "About"
         Height          =   375
         Left            =   120
         TabIndex        =   8
         Top             =   2040
         Width           =   1575
      End
      Begin VB.CommandButton cmdExit 
         Caption         =   "E&xit"
         Height          =   375
         Left            =   120
         TabIndex        =   7
         ToolTipText     =   "Close delete window"
         Top             =   3000
         Width           =   1575
      End
      Begin VB.CommandButton cmdShortcutPathDialog 
         Caption         =   "&Browse for files"
         Height          =   375
         Left            =   120
         TabIndex        =   6
         ToolTipText     =   "Browse for Directory"
         Top             =   1560
         Width           =   1575
      End
      Begin VB.PictureBox Picture3 
         AutoSize        =   -1  'True
         Height          =   540
         Left            =   1800
         Picture         =   "BJ's Delete Files.frx":08CA
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   4
         ToolTipText     =   "Click to Delete"
         Top             =   840
         Width           =   540
      End
      Begin VB.CommandButton Command2 
         Caption         =   "&Delete without Recycle bin"
         Default         =   -1  'True
         Height          =   495
         Left            =   120
         OLEDropMode     =   1  'Manual
         TabIndex        =   5
         ToolTipText     =   "Click to Delete"
         Top             =   840
         Width           =   1575
      End
      Begin VB.PictureBox Picture1 
         AutoSize        =   -1  'True
         Height          =   540
         Left            =   1800
         Picture         =   "BJ's Delete Files.frx":1194
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   3
         ToolTipText     =   "Send to Recycle Bin"
         Top             =   240
         Width           =   540
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Delete to the &Recycle Bin"
         Height          =   495
         Left            =   120
         OLEDropMode     =   1  'Manual
         TabIndex        =   2
         ToolTipText     =   "Send to Recycle Bin"
         Top             =   240
         Width           =   1575
      End
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   840
      Top             =   1560
   End
   Begin VB.FileListBox File1 
      Height          =   3405
      Left            =   0
      MultiSelect     =   2  'Extended
      OLEDragMode     =   1  'Automatic
      OLEDropMode     =   1  'Manual
      TabIndex        =   0
      ToolTipText     =   "Files that can be Deleted"
      Top             =   0
      Width           =   3735
   End
End
Attribute VB_Name = "frmDeleteFiles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
    (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, _
    ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Dim iGrabX As Integer
Dim iGrabY As Integer
Dim ControlZOrder As Long

'Function to get Windows directory
Private Declare Function GetWindowsDirectory Lib "kernel32" Alias _
"GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long

'Variable to store the Windows directory
Dim WinDir As String

'Buffer and constant used for API functions
Dim msBuffer As String * 255
Const BUFFERSIZE = 255

Private Type SHFILEOPSTRUCT
     hwnd As Long
     wFunc As Long
     pFrom As String
     pTo As String
     fFlags As Integer
     fAnyOperationsAborted As Boolean
     hNameMappings As Long
     lpszProgressTitle As String
End Type

Private Declare Function SHFileOperation Lib "shell32.dll" Alias "SHFileOperationA" (lpFileOp As SHFILEOPSTRUCT) As Long

Private Const FO_DELETE = &H3
Private Const FOF_ALLOWUNDO = &H40
Private Declare Function ShellAbout Lib "shell32.dll" Alias "ShellAboutA" _
(ByVal hwnd As Long, ByVal szApp As String, ByVal szOtherStuff As String, _
ByVal hIcon As Long) As Long



Public Sub AboutBox(hwnd As Long)
'-----------------------------------------------------------------
    ' Show help about dialog...
    ShellAbout hwnd, App.EXEName, _
               vbCrLf & "BJ. E-Mail me: bryce3@bigpond.com", frmDeleteFiles.Icon
'-----------------------------------------------------------------
End Sub

Private Sub cmdAbout_Click()
AboutBox Me.hwnd

End Sub

Private Sub cmdEMail_Click()
    ShellExecute 0, "Open", "mailto:bryce3@bigpond.com?subject=Info about BJ's How to Get... Delete Files.&body=Type here the message that you want to send me.", "", "", vbMaximizedFocus

End Sub

Private Sub cmdExit_Click()
If App.EXEName = "BJ's Delete Files" Then
End
Else
Unload frmDeleteFiles
frmHowtoGet.Show
End If

End Sub

Private Sub cmdShortcutPathDialog_Click()
Dim udtBrowseInfo As BROWSEINFO
Dim lRet As Long
Dim lPathID As Long
Dim sPath As String
Dim nNullPos As Integer

File1.SetFocus

'Specify the window handle for the owner of the dialog box
udtBrowseInfo.hOwner = Me.hwnd

'Specify the root to start browsing from;
'if null, My Computer is the root
udtBrowseInfo.pidlRoot = 0&

'Specify a title.  This is not the caption of the dialog.  Useful for
'adding any kind of additional information or instructions
udtBrowseInfo.lpszTitle = "Select a folder"

'Specify any flags; See Declarations section
udtBrowseInfo.ulFlags = BIF_RETURNONLYFSDIRS

'Call the function.
'The return value is a pointer to an item identifier list that
'specifies the location of the selected folder.
'If the user cancels the dialog box, the return value is 0.
lPathID = SHBrowseForFolder(udtBrowseInfo)

sPath = Space$(512)
lRet = SHGetPathFromIDList(lPathID, sPath)

If lRet Then
    nNullPos = InStr(sPath, vbNullChar)
    File1 = Left(sPath, nNullPos - 1)
End If

End Sub

Private Sub Command1_Click()
Dim FileOperation As SHFILEOPSTRUCT
Dim lReturn As Long

If File1.ListIndex = -1 Then
    MsgBox "Are you sure you want to delete this file", vbOKCancel + vbQuestion, "Delete"
    File1.SetFocus
End If
    If vbOK Then
    Picture1.Picture = Picture2.Picture
    frmDeleteFiles.Icon = Picture2.Picture
    Else
    If vbCancel Then
    Exit Sub
End If
End If


With FileOperation
    .wFunc = FO_DELETE
    .pFrom = File1.Path & "\" & File1.List(File1.ListIndex)     'fichier sélectionné dans la liste
    .fFlags = FOF_ALLOWUNDO
End With

lReturn = SHFileOperation(FileOperation)

Timer1.Enabled = True

End Sub

Private Sub Command1_DragDrop(Source As Control, X As Single, Y As Single)
Command1_Click
Timer1.Enabled = True

End Sub

Private Sub Command2_Click()
Kill File1.Path & "\" & File1.List(File1.ListIndex)
Timer1.Enabled = True
End Sub

Private Sub Form_Load()
If App.PrevInstance = True Then
End
End If
CreateKey "HKEY_CURRENT_USER\Software\BJ\BJ's How to Get..."
SetStringValue "HKEY_CURRENT_USER\Software\BJ\BJ's How to Get...", "Delete Files", App.EXEName & "- Last time run was " & " - " & Now
frmDeleteFiles.Icon = Picture1
End Sub

Private Sub Picture1_DblClick()
Dim FileOperation As SHFILEOPSTRUCT
    FileOperation.fFlags = FOF_ALLOWUNDO

End Sub

Private Sub Picture1_DragDrop(Source As Control, X As Single, Y As Single)
Command1_Click
Timer1.Enabled = True

End Sub

Private Sub File1_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
'control was dropped somewhere so move it to the point where it was dropped and offset it by the coordinates within the control where you are dragging
File1.Move File1.Left + X - iGrabX, File1.Top + Y - iGrabY
End Sub
Private Sub File1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbLeftButton Then
    'remember what part of the control you are dragging by
    iGrabX = X
    iGrabY = Y
    
    'begin dragging the control
    File1.Drag vbBeginDrag
Else
    ControlZOrder = File1.hwnd
End If
End Sub
Private Sub File1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbLeftButton Then
    'mouse button released so stop dragging
    File1.Drag vbEndDrag
End If
End Sub

Private Sub Picture3_Click()
Command2_Click
Timer1.Enabled = True
End Sub

Private Sub Timer1_Timer()
      File1.Refresh
      Timer1.Enabled = False
End Sub
