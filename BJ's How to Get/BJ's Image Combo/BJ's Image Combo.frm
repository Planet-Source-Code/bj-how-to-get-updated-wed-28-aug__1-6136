VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmImageCombo 
   BackColor       =   &H80000004&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "BJ's How to Get... Image Combo."
   ClientHeight    =   945
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5220
   Icon            =   "BJ's Image Combo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   945
   ScaleWidth      =   5220
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   2640
      TabIndex        =   6
      Top             =   320
      Visible         =   0   'False
      Width           =   1425
      _ExtentX        =   2514
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
      Max             =   10
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   2760
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   19
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "BJ's Image Combo.frx":014A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "BJ's Image Combo.frx":0464
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "BJ's Image Combo.frx":077E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "BJ's Image Combo.frx":0A98
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "BJ's Image Combo.frx":0DB2
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "BJ's Image Combo.frx":10CC
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "BJ's Image Combo.frx":13E6
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "BJ's Image Combo.frx":1700
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "BJ's Image Combo.frx":1A1A
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "BJ's Image Combo.frx":1D34
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "BJ's Image Combo.frx":204E
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "BJ's Image Combo.frx":2368
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "BJ's Image Combo.frx":2682
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "BJ's Image Combo.frx":299C
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "BJ's Image Combo.frx":2CB6
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "BJ's Image Combo.frx":2FD0
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "BJ's Image Combo.frx":32EA
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "BJ's Image Combo.frx":3884
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "BJ's Image Combo.frx":3CD6
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Left            =   3600
      Top             =   120
   End
   Begin VB.TextBox Color 
      Height          =   890
      Left            =   2520
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   0
      Width           =   2655
   End
   Begin MSComctlLib.ImageCombo ImageCombo1 
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   661
      _Version        =   393216
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Text            =   "Select a Colour"
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   3480
      TabIndex        =   3
      Text            =   "00"
      Top             =   1320
      Width           =   375
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   3000
      TabIndex        =   4
      Text            =   "00"
      Top             =   1320
      Width           =   375
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   3960
      TabIndex        =   5
      Text            =   "00"
      Top             =   1320
      Width           =   375
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   615
      Left            =   2400
      TabIndex        =   2
      Top             =   3480
      Width           =   2415
   End
End
Attribute VB_Name = "frmImageCombo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
    (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, _
    ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Dim objNewItem As ComboItem
Dim Hours As Integer
Dim Minutes As Integer
Dim Seconds As Integer
Dim Time As Date
Private Declare Function ShellAbout Lib "shell32.dll" Alias "ShellAboutA" _
(ByVal hwnd As Long, ByVal szApp As String, ByVal szOtherStuff As String, _
ByVal hIcon As Long) As Long



Public Sub AboutBox(hwnd As Long)
'-----------------------------------------------------------------
    ' Show help about dialog...
    ShellAbout hwnd, App.EXEName, _
               vbCrLf & "BJ. E-Mail me: bryce3@bigpond.com", frmImageCombo.Icon
'-----------------------------------------------------------------
End Sub

Private Sub Form_Click()
 Text3.Text = 10
    Timer1.Interval = 50
    Seconds = 10
    Time = 0
Label1.Caption = "Closing in 02 Seconds"
Color.Text = "You have chosen to Exit" & vbNewLine
Mydisplay
Timer1.Enabled = True
End Sub

Private Sub ImageCombo1_Click()
If ImageCombo1.Text = "Exit" Then
Mydisplay
Timer1.Enabled = True
ElseIf ImageCombo1.Text = "E-Mail" Then
ShellExecute 0, "Open", "mailto:bryce3@bigpond.com?subject=Info about BJ's How to Get... Image Combo.&body=Type here the message that you want to send me.", "", "", vbMaximizedFocus
ElseIf ImageCombo1.Text = "About" Then
AboutBox Me.hwnd
ElseIf ImageCombo1.Text = "Blue" Then '1
Color.BackColor = &HFF0000    '1
ElseIf ImageCombo1.Text = "Dark Blue" Then '2
Color.BackColor = &H800000    '2
Color.ForeColor = &H80000004
ElseIf ImageCombo1.Text = "Light Blue" Then '3
Color.BackColor = &HFFFF00    '3
Color.ForeColor = &H0&
ElseIf ImageCombo1.Text = "Desktop" Then '4
Color.BackColor = &H80000001  '4
ElseIf ImageCombo1.Text = "Green" Then '5
Color.BackColor = &HFF00&     '5
Color.ForeColor = &H0&
ElseIf ImageCombo1.Text = "Dark Green" Then '6
Color.BackColor = &H8000&     '6
Color.ForeColor = &H80000004
ElseIf ImageCombo1.Text = "Yellow" Then '7
Color.BackColor = &HFFFF&     '7
Color.ForeColor = &H0&
ElseIf ImageCombo1.Text = "Dark Yellow" Then '8
Color.BackColor = &H8080&     '8
Color.ForeColor = &H80000004
ElseIf ImageCombo1.Text = "Red" Then '9
Color.BackColor = &HFF&           '9
ElseIf ImageCombo1.Text = "Dark Red" Then '10
Color.BackColor = &H80&           '10
Color.ForeColor = &H80000004
ElseIf ImageCombo1.Text = "Grey" Then '11
Color.BackColor = &H80000004      '11
Color.ForeColor = &H0&
ElseIf ImageCombo1.Text = "Dark Grey" Then '12
Color.BackColor = &H80000003      '12
ElseIf ImageCombo1.Text = "Magenta" Then '13
Color.BackColor = &HFF00FF        '13
ElseIf ImageCombo1.Text = "Dark Magenta" Then '14
Color.BackColor = &H800080        '14
Color.ForeColor = &H80000004
ElseIf ImageCombo1.Text = "Black" Then '15
Color.BackColor = &H80000007      '15
Color.ForeColor = &H80000004
ElseIf ImageCombo1.Text = "White" Then '16
Color.BackColor = &H80000005      '16
Color.ForeColor = &H0&
End If

Color.Text = "You have clicked on " & ImageCombo1.Text & vbNewLine
End Sub

Private Sub Form_Load()
If App.PrevInstance = True Then
End
End If
CreateKey "HKEY_CURRENT_USER\Software\BJ\BJ's How to Get..."
SetStringValue "HKEY_CURRENT_USER\Software\BJ\BJ's How to Get...", "Image Combo", App.EXEName & "- Last time run was " & " - " & Now
 Text3.Text = 10
    Timer1.Interval = 1000
    Seconds = 10
    Time = 0
   Set ImageCombo1.ImageList = ImageList1

   Set objNewItem = ImageCombo1.ComboItems.Add(1, _
   "Blue", "Blue")
   Set objNewItem = ImageCombo1.ComboItems.Add(2, _
   "Dark Blue", "Dark Blue")
   Set objNewItem = ImageCombo1.ComboItems.Add(3, _
   "Light Blue", "Light Blue")
   Set objNewItem = ImageCombo1.ComboItems.Add(4, _
   "Desktop", "Desktop")
   Set objNewItem = ImageCombo1.ComboItems.Add(5, _
   "Green", "Green")
   Set objNewItem = ImageCombo1.ComboItems.Add(6, _
   "Dark Green", "Dark Green")
   Set objNewItem = ImageCombo1.ComboItems.Add(7, _
   "Yellow", "Yellow")
   Set objNewItem = ImageCombo1.ComboItems.Add(8, _
   "Dark Yellow", "Dark Yellow")
   Set objNewItem = ImageCombo1.ComboItems.Add(9, _
   "Red", "Red")
   Set objNewItem = ImageCombo1.ComboItems.Add(10, _
   "Dark Red", "Dark Red")
   Set objNewItem = ImageCombo1.ComboItems.Add(11, _
   "Grey", "Grey")
   Set objNewItem = ImageCombo1.ComboItems.Add(12, _
   "Dark Grey", "Dark Grey")
   Set objNewItem = ImageCombo1.ComboItems.Add(13, _
   "Magenta", "Magenta")
   Set objNewItem = ImageCombo1.ComboItems.Add(14, _
   "Dark Magenta", "Dark Magenta")
   Set objNewItem = ImageCombo1.ComboItems.Add(15, _
   "Black", "Black")
   Set objNewItem = ImageCombo1.ComboItems.Add(16, _
   "White", "White")
   Set objNewItem = ImageCombo1.ComboItems.Add(17, _
   "About", "About")
   Set objNewItem = ImageCombo1.ComboItems.Add(18, _
   "E-Mail", "E-Mail")
   Set objNewItem = ImageCombo1.ComboItems.Add(19, _
   "Exit", "Exit")

ImageCombo1.ComboItems("Blue").Image = 1
ImageCombo1.ComboItems("Dark Blue").Image = 2
ImageCombo1.ComboItems("Light Blue").Image = 3
ImageCombo1.ComboItems("Desktop").Image = 4
ImageCombo1.ComboItems("Green").Image = 5
ImageCombo1.ComboItems("Dark Green").Image = 6
ImageCombo1.ComboItems("Yellow").Image = 7
ImageCombo1.ComboItems("Dark Yellow").Image = 8
ImageCombo1.ComboItems("Red").Image = 9
ImageCombo1.ComboItems("Dark Red").Image = 10
ImageCombo1.ComboItems("Grey").Image = 11
ImageCombo1.ComboItems("Dark Grey").Image = 12
ImageCombo1.ComboItems("Magenta").Image = 13
ImageCombo1.ComboItems("Dark Magenta").Image = 14
ImageCombo1.ComboItems("Black").Image = 15
ImageCombo1.ComboItems("White").Image = 16
ImageCombo1.ComboItems("About").Image = 17
ImageCombo1.ComboItems("E-Mail").Image = 18
ImageCombo1.ComboItems("Exit").Image = 19



ImageCombo1.Indentation = 1

Color.Text = frmImageCombo.Caption
End Sub

Private Sub Mydisplay()
    Seconds = Val(Text3.Text)
    
    Time = TimeSerial(Hours, Minutes, Seconds)
    
'    Label1.Caption = "Closing in " & Format$(Time, "hh") & ":" & Format$(Time, "nn") & ":" & Format$(Time, "ss") & " Seconds"
   If Label1.Caption = "Closing in 02 Seconds" Then
        Label1.Caption = "Closing in " & Format$(Time, "ss") & " Second"
Color.Text = "You have clicked on " & ImageCombo1.Text & vbNewLine & vbNewLine & vbNewLine & Label1.Caption
Else
Label1.Caption = "Closing in " & Format$(Time, "ss") & " Seconds"
Color.Text = "You have clicked on " & ImageCombo1.Text & vbNewLine & vbNewLine & vbNewLine & Label1.Caption
End If
End Sub

Private Sub Timer1_Timer()
  ProgressBar1.Visible = True
    Timer1.Enabled = False
    If (Format$(Time, "ss")) <> "00" Then  'Counter to continue loop until 0
        
        Time = DateAdd("s", -1, Time)
        Label1.Visible = False
'        Label1.Caption = "Closing in " & Format$(Time, "ss") & " Seconds"
        
        Label1.Visible = True
        Timer1.Enabled = True
    Else
        
        Timer1.Enabled = False
        Beep
If App.EXEName = "BJ's Image Combo" Then
End
Else
Unload frmImageCombo
frmHowtoGet.Show
End If
    End If
   If Label1.Caption = "Closing in 02 Seconds" Then
        Label1.Caption = "Closing in " & Format$(Time, "ss") & " Second"
Color.Text = "You have chosen to Exit" & vbNewLine & vbNewLine & vbNewLine & Label1.Caption
Else
Label1.Caption = "Closing in " & Format$(Time, "ss") & " Seconds"
Color.Text = "You have chosen to Exit" & vbNewLine & vbNewLine & vbNewLine & Label1.Caption
End If
ProgressBar1.Value = Format$(Time, "ss")
End Sub

