VERSION 5.00
Begin VB.Form frmRenameFiles 
   Appearance      =   0  'Flat
   Caption         =   "BJ's Rename Files"
   ClientHeight    =   5535
   ClientLeft      =   1095
   ClientTop       =   1485
   ClientWidth     =   8745
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000008&
   Icon            =   "BJ's Rename Files.frx":0000
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5535
   ScaleWidth      =   8745
   Begin VB.CommandButton Command5 
      Cancel          =   -1  'True
      Caption         =   "E-Mail"
      Height          =   375
      Left            =   5520
      TabIndex        =   15
      Top             =   1320
      Width           =   1455
   End
   Begin VB.CommandButton Command4 
      Caption         =   "About"
      Height          =   375
      Left            =   3840
      TabIndex        =   14
      Top             =   1320
      Width           =   1455
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   1920
      TabIndex        =   12
      Text            =   "File1"
      Top             =   2400
      Width           =   5175
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Rename Now"
      Height          =   735
      Left            =   7200
      TabIndex        =   11
      Top             =   1920
      Width           =   1455
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   1920
      TabIndex        =   9
      Text            =   "File1"
      Top             =   1920
      Width           =   5175
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Exit"
      Height          =   375
      Left            =   7200
      TabIndex        =   8
      Top             =   1320
      Width           =   1455
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   6360
      Top             =   1320
   End
   Begin VB.FileListBox File1 
      Height          =   2430
      Left            =   4440
      TabIndex        =   7
      Top             =   3000
      Width           =   4215
   End
   Begin VB.DirListBox Dir1 
      Height          =   2115
      Left            =   120
      TabIndex        =   6
      Top             =   3360
      Width           =   4215
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   120
      TabIndex        =   5
      Top             =   3000
      Width           =   4215
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   1920
      TabIndex        =   3
      Text            =   "File1"
      Top             =   720
      Width           =   5175
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Rename Now"
      Height          =   735
      Left            =   7200
      TabIndex        =   2
      Top             =   240
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1920
      Locked          =   -1  'True
      TabIndex        =   1
      Text            =   "C:\Temp"
      Top             =   240
      Width           =   5175
   End
   Begin VB.Label Label4 
      Caption         =   "Rename 1 File to:"
      Height          =   375
      Left            =   120
      TabIndex        =   13
      Top             =   2400
      Width           =   1815
   End
   Begin VB.Label Label3 
      Caption         =   "Rename 1 File from:"
      Height          =   375
      Left            =   120
      TabIndex        =   10
      Top             =   1920
      Width           =   1815
   End
   Begin VB.Label Label2 
      Caption         =   "Rename all Files to:"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   720
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "Rename all Files in:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   1695
   End
End
Attribute VB_Name = "frmRenameFiles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
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
               vbCrLf & "BJ. E-Mail me: bryce3@bigpond.com", frmRenameFiles.Icon
'-----------------------------------------------------------------
End Sub


Private Sub Command1_Click()
   RenameFiles Text1.Text
Timer1.Enabled = True
End Sub
Private Function GiveNewName%(ByVal Oldname$, ByVal NewName$)

' Description
'     Renames OldName$ to NewName$
'
' Parameters
'     Name           Type        Value
'     --------------------------------------------------
'     OldName$       String      The file to be renamed
'     NewName$       String      The new filename
'
'  Returns
'     True if everything went OK
'     False if there was an error
'

   ' Trap any errors
   On Error GoTo Error_In_Renaming

   ' Now rename the file
   Name Oldname$ As NewName$

   ' Return True to indicate a success
   GiveNewName% = True

Exit_The_Function:
   
   Exit Function

Error_In_Renaming:

   ' Return False to indicate failure
   MsgBox Error$(Err)
   GiveNewName% = False
   Resume Exit_The_Function

End Function

Private Sub RenameFiles(Path$)

' Description
'     Renames all files in the specified Path
'
' Parameters
'     Name           Type     Value
'     -------------------------------------------------------------------
'     Path$          String   The path where the files to be renamed are
'
' Returns
'     Nothing
'


Dim t$      ' To hold the filename while renaming
Dim ok%     ' To hold the return values from GiveNewName% and MsgBox
Dim Counter%   ' Counts the files

   ' Get the first filename
   t$ = Dir$(Path$ & "\*.*")
   
   ' Keep on renaming until the directory contains no unrenamed files
   Do While t$ <> ""

      ' Increment the counter
      Counter% = Counter% + 1
      
      ' Give new filename that will look like this "FILE####.$$$" where "####" is the Counter%
      ' value and the "$$$" is the original file's extension
      ok% = GiveNewName%(Path$ & "\" & t$, Path$ & "\" & Text2.Text & CStr(Counter%) & Right$(t$, 4))

      ' If there was an error renaming the file...
      If Not ok% Then
         ' Ask the user to proceed or quit
         ok% = MsgBox("There was an error renaming " & t$ & " to " & "File" & CStr(Counter%) & Right$(t$, 4) & Chr$(13) & Chr$(13) & "Proceed?", 4 + 32, "Rename files")
         If ok% = 7 Then   ' No was selected
            ' Quit
            End
         End If
      End If

      ' Get new filename
      t$ = Dir$

   Loop

End Sub

Private Sub Command2_Click()
End
End Sub

Private Sub Command3_Click()
Dim t$      ' To hold the filename while renaming
Dim ok%     ' To hold the return values from GiveNewName% and MsgBox
Dim Counter%   ' Counts the files

   ' Get the first filename
   t$ = Text3.Text
   
   ' Keep on renaming until the directory contains no unrenamed files
   Do While t$ <> ""

      ' Increment the counter
      
      ' Give new filename that will look like this "FILE####.$$$" where "####" is the Counter%
      ' value and the "$$$" is the original file's extension
      ok% = GiveNewName%(Text4.Text & CStr(Counter%) & Right$(t$, 4))

      ' If there was an error renaming the file...
      If Not ok% Then
         ' Ask the user to proceed or quit
         ok% = MsgBox("There was an error renaming " & " to " & Text4.Text & CStr(Counter%) & Right$(t$, 4) & Chr$(13) & Chr$(13) & "Proceed?", 4 + 32, "Rename files")
         If ok% = 7 Then   ' No was selected
            ' Quit
            End
         End If
      End If

      ' Get new filename
      t$ = Text3.Text

   Loop

End Sub

Private Sub Command4_Click()
AboutBox Me.hwnd
End Sub

Private Sub Command5_Click()
    ShellExecute 0, "Open", "mailto:bryce3@bigpond.com?subject=Info about BJ's How to Get... Rename Files.&body=Type here the message that you want to send me.", "", "", vbMaximizedFocus

End Sub

Private Sub Dir1_Change()
    File1.Path = Dir1.Path
    Text1.Text = Dir1.Path
End Sub

Private Sub Drive1_Change()
    Dir1.Path = Drive1.Drive
End Sub

Private Sub File1_DblClick()
Dim i As Integer
Text3.Text = File1.FileName
End Sub
Private Sub Form_Load()
If App.PrevInstance = True Then
End
End If
CreateKey "HKEY_CURRENT_USER\Software\BJ\BJ's How to Get..."
SetStringValue "HKEY_CURRENT_USER\Software\BJ\BJ's How to Get...", "Rename Files", App.EXEName & "- Last time run was " & " - " & Now

    Text1.Text = Dir1.Path
End Sub
Private Sub Text1_Click()
MsgBox "Sorry, you can't change from here." & vbCrLf & _
"To change Directory click on the drive and path below.", vbInformation + vbOKOnly, "Error"
End Sub

Private Sub Timer1_Timer()
File1.Refresh
Timer1.Enabled = False
End Sub
