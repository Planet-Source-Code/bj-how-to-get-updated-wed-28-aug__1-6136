VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form frmBasicCalender 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "BJ's Basic Calender - "
   ClientHeight    =   5550
   ClientLeft      =   2370
   ClientTop       =   465
   ClientWidth     =   7635
   Icon            =   "BJ's Basic Calender.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5550
   ScaleWidth      =   7635
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   8760
      Top             =   2760
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.TextBox txtLinkName 
      Height          =   285
      Left            =   8400
      TabIndex        =   43
      Top             =   480
      Width           =   120
   End
   Begin VB.TextBox txtExeName 
      Height          =   285
      Left            =   8400
      TabIndex        =   42
      Top             =   855
      Width           =   120
   End
   Begin VB.TextBox txtWorkDir 
      Height          =   285
      Left            =   8400
      TabIndex        =   41
      Top             =   1245
      Width           =   120
   End
   Begin VB.TextBox txtIconFile 
      Height          =   285
      Left            =   8400
      TabIndex        =   40
      Top             =   2730
      Width           =   120
   End
   Begin VB.TextBox txtIconIndex 
      Height          =   285
      Left            =   8400
      TabIndex        =   39
      Top             =   3105
      Width           =   120
   End
   Begin VB.TextBox txtCmdArgs 
      Height          =   285
      Left            =   8400
      TabIndex        =   38
      Top             =   1620
      Width           =   120
   End
   Begin VB.TextBox txtShowCmd 
      Height          =   285
      Left            =   8400
      TabIndex        =   37
      Top             =   3495
      Width           =   120
   End
   Begin VB.TextBox txtlnk 
      Height          =   285
      Left            =   8400
      TabIndex        =   36
      Top             =   3900
      Width           =   120
   End
   Begin VB.TextBox txtLnkDesc 
      Height          =   285
      Left            =   8400
      TabIndex        =   35
      Top             =   2040
      Width           =   150
   End
   Begin VB.TextBox txtwHotKey 
      Height          =   285
      Left            =   8400
      TabIndex        =   34
      Top             =   2400
      Width           =   150
   End
   Begin VB.Timer tmrCheckTime 
      Interval        =   1000
      Left            =   8880
      Top             =   2280
   End
   Begin VB.Timer popupTmr 
      Interval        =   1000
      Left            =   8880
      Top             =   1800
   End
   Begin VB.Frame Frame4 
      Height          =   1335
      Left            =   40
      TabIndex        =   28
      ToolTipText     =   "Right Click for Popup Menu"
      Top             =   4160
      Width           =   7530
      Begin VB.Line Line13 
         BorderColor     =   &H80000005&
         X1              =   5400
         X2              =   5400
         Y1              =   840
         Y2              =   1200
      End
      Begin VB.Label lblDays 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "888 Days"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   120
         TabIndex        =   33
         ToolTipText     =   "Right click for popup menu."
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label lblHrs 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "88 Hours"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   1680
         TabIndex        =   32
         ToolTipText     =   "Right click for popup menu."
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label lblMins 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "88 Minutes"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   3480
         TabIndex        =   31
         ToolTipText     =   "Right click for popup menu."
         Top             =   840
         Width           =   1695
      End
      Begin VB.Label lblSecs 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "88 Seconds"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5640
         TabIndex        =   30
         ToolTipText     =   "Right click for popup menu."
         Top             =   840
         Width           =   1695
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000005&
         Index           =   5
         X1              =   3240
         X2              =   3240
         Y1              =   1200
         Y2              =   840
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000005&
         Index           =   7
         X1              =   1560
         X2              =   1560
         Y1              =   1200
         Y2              =   840
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000005&
         Index           =   0
         X1              =   120
         X2              =   7440
         Y1              =   720
         Y2              =   720
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Your Computer has been on now for..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   435
         Left            =   510
         TabIndex        =   29
         ToolTipText     =   "Right Click for Popup Menu"
         Top             =   180
         Width           =   6645
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000010&
         BorderWidth     =   2
         Index           =   8
         X1              =   120
         X2              =   7440
         Y1              =   720
         Y2              =   720
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000010&
         BorderWidth     =   2
         Index           =   4
         X1              =   3240
         X2              =   3240
         Y1              =   1200
         Y2              =   840
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000010&
         BorderWidth     =   2
         Index           =   6
         X1              =   1560
         X2              =   1560
         Y1              =   1200
         Y2              =   840
      End
      Begin VB.Line Line14 
         BorderColor     =   &H80000010&
         BorderWidth     =   2
         X1              =   5400
         X2              =   5400
         Y1              =   840
         Y2              =   1200
      End
   End
   Begin VB.Frame Frame3 
      Height          =   735
      Left            =   40
      TabIndex        =   26
      ToolTipText     =   "Double Click to Change Time Zone. Middle Button for Colours. Right Click for Popup Menu"
      Top             =   3420
      Width           =   7530
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "(GMT+-88:88) Abcdefghijklmnopqrstuvwxyz"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   27
         ToolTipText     =   "Double Click to Change Time Zone. Right Click for Popup Menu"
         Top             =   240
         Width           =   7335
      End
   End
   Begin VB.Frame Frame2 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   1.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3375
      Left            =   3840
      TabIndex        =   5
      ToolTipText     =   "Double Click to Change Date & Time. Right Click for Popup Menu"
      Top             =   40
      Width           =   3735
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   240
         ScaleHeight     =   33
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   33
         TabIndex        =   46
         ToolTipText     =   "Double Click to Change Date & Time. Right Click for Popup Menu"
         Top             =   120
         Width           =   495
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Weeks left:             88"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   390
         Left            =   195
         TabIndex        =   45
         ToolTipText     =   "Double Click to Change Date & Time. Right Click for Popup Menu"
         Top             =   2880
         Width           =   3255
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Days left:             888"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   390
         Left            =   240
         TabIndex        =   44
         ToolTipText     =   "Double Click to Change Date & Time. Right Click for Popup Menu"
         Top             =   1920
         Width           =   3195
      End
      Begin VB.Label lblDisClock 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Â"
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   27.75
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   615
         Left            =   120
         TabIndex        =   24
         ToolTipText     =   "Click to Connect to Atomic Clock. (Must have Dial up Connection)"
         Top             =   840
         Width           =   555
      End
      Begin VB.Line Line12 
         BorderColor     =   &H80000005&
         X1              =   120
         X2              =   3600
         Y1              =   2460
         Y2              =   2460
      End
      Begin VB.Line Line11 
         BorderColor     =   &H80000005&
         X1              =   120
         X2              =   3600
         Y1              =   1560
         Y2              =   1560
      End
      Begin VB.Line Line10 
         BorderColor     =   &H80000005&
         X1              =   120
         X2              =   3600
         Y1              =   720
         Y2              =   720
      End
      Begin VB.Line Line6 
         BorderColor     =   &H80000010&
         BorderWidth     =   2
         X1              =   120
         X2              =   3600
         Y1              =   2460
         Y2              =   2460
      End
      Begin VB.Line Line5 
         BorderColor     =   &H80000010&
         BorderWidth     =   2
         X1              =   120
         X2              =   3600
         Y1              =   1560
         Y2              =   1560
      End
      Begin VB.Line Line4 
         BorderColor     =   &H80000010&
         BorderWidth     =   2
         X1              =   120
         X2              =   3600
         Y1              =   720
         Y2              =   720
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Starsigns"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   21.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   1200
         TabIndex        =   11
         ToolTipText     =   "Double Click to Change Date & Time. Right Click for Popup Menu"
         Top             =   0
         Width           =   1725
      End
      Begin VB.Label lblWeekofYear 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "88 of 88"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   390
         Left            =   2240
         TabIndex        =   10
         ToolTipText     =   "Double Click to Change Date & Time. Right Click for Popup Menu"
         Top             =   2520
         Width           =   1215
      End
      Begin VB.Label lblDayofYear 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "888 of 888"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   390
         Left            =   1840
         TabIndex        =   9
         ToolTipText     =   "Double Click to Change Date & Time. Right Click for Popup Menu"
         Top             =   1560
         Width           =   1575
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Week"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   390
         Left            =   240
         TabIndex        =   8
         ToolTipText     =   "Double Click to Change Date & Time. Right Click for Popup Menu"
         Top             =   2520
         Width           =   885
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Day"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   390
         Left            =   260
         TabIndex        =   7
         ToolTipText     =   "Double Click to Change Date & Time. Right Click for Popup Menu"
         Top             =   1560
         Width           =   645
      End
      Begin VB.Label lblTime 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "88:88:88 ampm"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Left            =   510
         TabIndex        =   6
         ToolTipText     =   "Double Click to Change Date & Time. Right Click for Popup Menu"
         Top             =   840
         Width           =   3225
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "September 24 to October 23"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   885
         TabIndex        =   47
         Top             =   480
         Width           =   2445
      End
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   1.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3375
      Left            =   40
      TabIndex        =   0
      ToolTipText     =   "Double Click to Change Date & Time. Right Click for Popup Menu"
      Top             =   40
      Width           =   3735
      Begin VB.Line Line9 
         BorderColor     =   &H80000005&
         X1              =   120
         X2              =   3480
         Y1              =   2700
         Y2              =   2700
      End
      Begin VB.Line Line8 
         BorderColor     =   &H80000005&
         X1              =   120
         X2              =   3480
         Y1              =   2040
         Y2              =   2040
      End
      Begin VB.Line Line7 
         BorderColor     =   &H80000005&
         X1              =   120
         X2              =   3480
         Y1              =   720
         Y2              =   720
      End
      Begin VB.Line Line3 
         BorderColor     =   &H80000010&
         BorderWidth     =   2
         X1              =   120
         X2              =   3480
         Y1              =   2700
         Y2              =   2700
      End
      Begin VB.Line Line2 
         BorderColor     =   &H80000010&
         BorderWidth     =   2
         X1              =   120
         X2              =   3480
         Y1              =   2040
         Y2              =   2040
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000010&
         BorderWidth     =   2
         Index           =   1
         X1              =   120
         X2              =   3480
         Y1              =   720
         Y2              =   720
      End
      Begin VB.Label lblMonth 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "September"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   540
         Left            =   735
         TabIndex        =   4
         ToolTipText     =   "Double Click to Change Date & Time. Right Click for Popup Menu"
         Top             =   2040
         Width           =   2235
      End
      Begin VB.Label lblDay 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H0000FF00&
         BackStyle       =   0  'Transparent
         Caption         =   "Wednesday"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Left            =   120
         TabIndex        =   3
         ToolTipText     =   "Double Click to Change Date & Time. Right Click for Popup Menu"
         Top             =   60
         Width           =   3465
      End
      Begin VB.Label lblYear 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "8888"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Left            =   1335
         TabIndex        =   2
         ToolTipText     =   "Double Click to Change Date & Time. Right Click for Popup Menu"
         Top             =   2760
         Width           =   975
      End
      Begin VB.Label lblNumber 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H000000C0&
         BackStyle       =   0  'Transparent
         Caption         =   "88"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   72
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1665
         Left            =   1080
         TabIndex        =   1
         ToolTipText     =   "Double Click to Change Date & Time. Right Click for Popup Menu"
         Top             =   600
         Width           =   1605
      End
   End
   Begin VB.Timer timDisplay 
      Interval        =   1
      Left            =   3000
      Top             =   240
   End
   Begin VB.Timer Timer1 
      Interval        =   970
      Left            =   2880
      Top             =   960
   End
   Begin VB.Label Label5 
      Caption         =   "Label5"
      Height          =   255
      Left            =   8400
      TabIndex        =   25
      Top             =   4320
      Width           =   735
   End
   Begin VB.Label lblPic 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Á"
      BeginProperty Font 
         Name            =   "Wingdings"
         Size            =   21.75
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Index           =   11
      Left            =   2640
      TabIndex        =   23
      Top             =   1200
      Visible         =   0   'False
      Width           =   390
   End
   Begin VB.Label lblPic 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "À"
      BeginProperty Font 
         Name            =   "Wingdings"
         Size            =   21.75
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Index           =   10
      Left            =   2280
      TabIndex        =   22
      Top             =   1200
      Visible         =   0   'False
      Width           =   390
   End
   Begin VB.Label lblPic 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "¿"
      BeginProperty Font 
         Name            =   "Wingdings"
         Size            =   21.75
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Index           =   9
      Left            =   1920
      TabIndex        =   21
      Top             =   1200
      Visible         =   0   'False
      Width           =   390
   End
   Begin VB.Label lblPic 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "¾"
      BeginProperty Font 
         Name            =   "Wingdings"
         Size            =   21.75
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Index           =   8
      Left            =   1560
      TabIndex        =   20
      Top             =   1200
      Visible         =   0   'False
      Width           =   390
   End
   Begin VB.Label lblPic 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "½"
      BeginProperty Font 
         Name            =   "Wingdings"
         Size            =   21.75
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Index           =   7
      Left            =   1200
      TabIndex        =   19
      Top             =   1200
      Visible         =   0   'False
      Width           =   390
   End
   Begin VB.Label lblPic 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "¼"
      BeginProperty Font 
         Name            =   "Wingdings"
         Size            =   21.75
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Index           =   6
      Left            =   840
      TabIndex        =   18
      Top             =   1200
      Visible         =   0   'False
      Width           =   390
   End
   Begin VB.Label lblPic 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "»"
      BeginProperty Font 
         Name            =   "Wingdings"
         Size            =   21.75
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Index           =   5
      Left            =   2640
      TabIndex        =   17
      Top             =   840
      Visible         =   0   'False
      Width           =   390
   End
   Begin VB.Label lblPic 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "º"
      BeginProperty Font 
         Name            =   "Wingdings"
         Size            =   21.75
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Index           =   4
      Left            =   2280
      TabIndex        =   16
      Top             =   840
      Visible         =   0   'False
      Width           =   390
   End
   Begin VB.Label lblPic 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "¹"
      BeginProperty Font 
         Name            =   "Wingdings"
         Size            =   21.75
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Index           =   3
      Left            =   1920
      TabIndex        =   15
      Top             =   840
      Visible         =   0   'False
      Width           =   390
   End
   Begin VB.Label lblPic 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "¸"
      BeginProperty Font 
         Name            =   "Wingdings"
         Size            =   21.75
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Index           =   2
      Left            =   1560
      TabIndex        =   14
      Top             =   840
      Visible         =   0   'False
      Width           =   390
   End
   Begin VB.Label lblPic 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "·"
      BeginProperty Font 
         Name            =   "Wingdings"
         Size            =   21.75
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Index           =   1
      Left            =   1200
      TabIndex        =   13
      Top             =   840
      Visible         =   0   'False
      Width           =   390
   End
   Begin VB.Label lblPic 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Â"
      BeginProperty Font 
         Name            =   "Wingdings"
         Size            =   21.75
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Index           =   0
      Left            =   840
      TabIndex        =   12
      Top             =   840
      Visible         =   0   'False
      Width           =   390
   End
   Begin VB.Menu mnuPopup 
      Caption         =   "PopupMenu"
      Visible         =   0   'False
      Begin VB.Menu mnuDateTime 
         Caption         =   "Adjust Date / Time"
      End
      Begin VB.Menu mnuTimeZone 
         Caption         =   "Change Time Zone"
      End
      Begin VB.Menu mnubar0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAtomicTime 
         Caption         =   "Atomic Clock"
      End
      Begin VB.Menu mnubar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuColor 
         Caption         =   "Colour"
      End
      Begin VB.Menu mnuHour 
         Caption         =   "12 Hour Time"
      End
      Begin VB.Menu mnubar2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuShowHideTimeZone 
         Caption         =   "Show TimeZone"
      End
      Begin VB.Menu mnuHide 
         Caption         =   "Show"
      End
      Begin VB.Menu mnubar3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuWinStart 
         Caption         =   "Add to Startup"
      End
      Begin VB.Menu mnuRunMenu 
         Caption         =   "Add to Run Command"
      End
      Begin VB.Menu mnubar4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEMail 
         Caption         =   "E-Mail"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "About"
      End
      Begin VB.Menu mnuHelp 
         Caption         =   "Help"
         Visible         =   0   'False
      End
      Begin VB.Menu mnubar5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "frmBasicCalender"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'for systems about box
Private Declare Function ShellAbout Lib "shell32.dll" Alias "ShellAboutA" _
(ByVal hwnd As Long, ByVal szApp As String, ByVal szOtherStuff As String, _
ByVal hIcon As Long) As Long

'for animated cursor
Dim mhBaseCursor As Long, mhAniCursor As Long
Dim state As Integer
Dim lResult As Long

'get screen size
Dim screenx As Long
Dim screeny As Long

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

'Used to delete to Recycle Bin
Private Const FO_DELETE = &H3
Private Const FOF_ALLOWUNDO = &H40

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

'gets windows running time
Private Declare Function GetTickCount Lib "kernel32" () As Long

Const MS_PER_SEC As Long = 1000
Const MS_PER_MIN = MS_PER_SEC * 60
Const MS_PER_HR = MS_PER_MIN * 60
Const MS_PER_DAY = MS_PER_HR * 24

Dim ms As Long
Dim secs As Long
Dim mins As Long
Dim hrs As Long
Dim days As Long

Dim Today As Variant
Dim RotateMyClock As Integer 'use to rotate the image of the clock
Dim MyTime As String
Dim daynum

Private Const INTERNET_OPEN_TYPE_PRECONFIG = 0
Private Const INTERNET_OPEN_TYPE_DIRECT = 1
Private Const INTERNET_OPEN_TYPE_PROXY = 3

Private Const scUserAgent = "VB Project"
Private Const INTERNET_FLAG_RELOAD = &H80000000

'used to get the atomic time
Private Declare Function InternetOpen Lib "wininet.dll" _
  Alias "InternetOpenA" (ByVal sAgent As String, _
  ByVal lAccessType As Long, ByVal sProxyName As String, _
  ByVal sProxyBypass As String, ByVal lFlags As Long) As Long

Private Declare Function InternetOpenUrl Lib "wininet.dll" _
  Alias "InternetOpenUrlA" (ByVal hOpen As Long, _
  ByVal sUrl As String, ByVal sHeaders As String, _
  ByVal lLength As Long, ByVal lFlags As Long, _
  ByVal lContext As Long) As Long

Private Declare Function InternetReadFile Lib "wininet.dll" _
  (ByVal hFile As Long, ByVal sBuffer As String, _
   ByVal lNumBytesToRead As Long, lNumberOfBytesRead As Long) _
  As Integer

Private Declare Function InternetCloseHandle _
   Lib "wininet.dll" (ByVal hInet As Long) As Integer

Private Declare Function GetTimeZoneInformation& Lib "kernel32" _
   (lpTimeZoneInformation As TIME_ZONE_INFORMATION)


Private Type SYSTEMTIME
        wYear As Integer
        wMonth As Integer
        wDayOfWeek As Integer
        wDay As Integer
        wHour As Integer
        wMinute As Integer
        wSecond As Integer
        wMilliseconds As Integer
End Type

Private Type TIME_ZONE_INFORMATION
        Bias As Long
        StandardName As String * 64
        StandardDate As SYSTEMTIME
        StandardBias As Long
        DaylightName As String * 64
        DaylightDate As SYSTEMTIME
        DaylightBias As Long
End Type

Public m_StrAtomicTime As String 'Optional Global Variable
'so caller can read back result

'Used to remove [X] and other menu items
'To understand this better see My app on Planet Source Code (BJ's How to Get...)
Private Declare Function DeleteMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Private Declare Function GetSystemMenu Lib "user32" (ByVal hwnd As Long, ByVal bRevert As Long) As Long

Private Const MF_BYPOSITION = &H400&

Private ReadyToClose As Boolean

Private Sub RemoveMenus(frm As Form, remove_close As Boolean)
Dim hMenu As Long
Dim hMenu1 As Long
    
    ' Get the form's system menu handle.
    hMenu = GetSystemMenu(hwnd, False)
    'disables right click meun from icon and [X] on form
'    If remove_close Then DeleteMenu hMenu, 6, MF_BYPOSITION ' Removes Close
'    If remove_close Then DeleteMenu hMenu, 5, MF_BYPOSITION ' Removes Seperator between Maximize and Close
'    If remove_close Then DeleteMenu hMenu, 1, MF_BYPOSITION ' Removes Move
'    If remove_close Then DeleteMenu hMenu, 4, MF_BYPOSITION ' Removes Maximize but not the Button
'    If remove_close Then DeleteMenu hMenu, 3, MF_BYPOSITION ' Removes Minimize but not the Button
'    If remove_close Then DeleteMenu hMenu, 2, MF_BYPOSITION ' Removes
'    If remove_close Then DeleteMenu hMenu, 0, MF_BYPOSITION ' Removes Remove
    'Note: You must have them in the order.
    'or you can do it this way
    hMenu1 = DeleteMenu(hMenu, 6, MF_BYPOSITION) And hMenu = DeleteMenu(hMenu, 5, MF_BYPOSITION)
'This goes into Form_Load. I already have it in there
'    RemoveMenus Me, True
    End Sub

Function IsLeapYear(ByVal sYear As String) As Boolean
'Used to get the extra day in a leap year.
    If IsDate("02/29/" & sYear) Then
            IsLeapYear = True
    Else
            IsLeapYear = False
    End If
End Function

'Private Sub StartCursorAnimation()
    
'    mhAniCursor = LoadCursorFromFile(App.Path & "\globe.ani")
'    lResult = SetClassLong((Frame1.hwnd), GCL_HCURSOR, mhAniCursor)
'    lResult = SetClassLong((Frame2.hwnd), GCL_HCURSOR, mhAniCursor)
'    lResult = SetClassLong((Frame3.hwnd), GCL_HCURSOR, mhAniCursor)
'    lResult = SetClassLong((Frame4.hwnd), GCL_HCURSOR, mhAniCursor)
'    lResult = SetClassLong((hwnd), GCL_HCURSOR, mhAniCursor)
'    state = 1

'End Sub

'Private Sub StopCursorAnimation()
    
'    lResult = SetClassLong((Frame1.hwnd), GCL_HCURSOR, mhBaseCursor)
'    lResult = SetClassLong((Frame2.hwnd), GCL_HCURSOR, mhBaseCursor)
'    lResult = SetClassLong((Frame3.hwnd), GCL_HCURSOR, mhBaseCursor)
'    lResult = SetClassLong((Frame4.hwnd), GCL_HCURSOR, mhBaseCursor)
'    lResult = SetClassLong((hwnd), GCL_HCURSOR, mhBaseCursor)
'    lResult = DestroyCursor(mhAniCursor)
    
'    state = 0

'End Sub

Private Sub tIcon() 'Set the Icon in the System Tray

    'sets cbSize to the Length of TrayIcon
    TrayIcon.cbSize = Len(TrayIcon)
    ' Handle of the window used to handle messages - which is the this form
    TrayIcon.hwnd = Me.hwnd
    ' ID code of the icon
    TrayIcon.uId = vbNull
    ' Flags
    TrayIcon.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
    ' ID of the call back message
    TrayIcon.ucallbackMessage = WM_MOUSEMOVE
    ' The icon - sets the icon that should be used
    TrayIcon.hIcon = Me.Icon
    ' The Tooltip for the icon - sets the Tooltip that will be displayed
    
    TrayIcon.szTip = "Right Click for Menu, Double Click to Show" & Chr$(0)

    ' Add icon to the tray by calling the Shell_NotifyIcon API
    'NIM_ADD is a Constant - add icon to tray
    Call Shell_NotifyIcon(NIM_ADD, TrayIcon)
    
    ' Don't let application appear in the Windows task list
    App.TaskVisible = False

End Sub

Private Sub Form_Load()


'Used to stop opening more than 1 instance of this app.
If App.PrevInstance = True Then
MsgBox "BJ's Basic Calender is already Running" & vbNewLine & _
"in the System Tray next to the System Clock.", _
 vbInformation, App.Title
End
Else
End If
       RemoveMenus Me, True
 
Call tIcon
    
'Creates and Set String Values in the Registry
CreateKey "HKEY_CURRENT_USER\Software\BJ\BJ's How to Get...\BJ's Basic Calender"
SetStringValue "HKEY_CURRENT_USER\Software\BJ", "", "bryce3@bigpond.com"
SetStringValue "HKEY_CURRENT_USER\Software\BJ\BJ's How to Get...", "", "bryce3@bigpond.com"
SetStringValue "HKEY_CURRENT_USER\Software\BJ\BJ's How to Get...", "Basic Calender", App.EXEName
SetStringValue "HKEY_CURRENT_USER\Software\BJ\BJ's How to Get...\BJ's Basic Calender", "", "bryce3@bigpond.com"

If FileExists(GetStringValue("HKEY_CURRENT_USER\Software\BJ\BJ's How to Get...\BJ's Basic Calender", "Windows Startup Path and File")) = True Then
mnuWinStart.Caption = "Remove from Startup"
End If

If GetStringValue("HKEY_CURRENT_USER\Software\BJ\BJ's How to Get...\BJ's Basic Calender", "Windows Startup Path and File") = "Error" Then ', txtLinkName.Text & txtlnk.Text
mnuWinStart.Caption = "Add to Startup"
Else
If FileExists(GetStringValue("HKEY_CURRENT_USER\Software\BJ\BJ's How to Get...\BJ's Basic Calender", "Windows Startup Path and File")) = True Then
mnuWinStart.Caption = "Remove from Startup"
End If
End If
 
If GetStringValue("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\BJ'sBC.exe", "") = "Error" Then
mnuRunMenu.Caption = "Add to Run Command"
Else
mnuRunMenu.Caption = "Remove from Run Command"
End If

If GetBinaryValue("HKEY_CURRENT_USER\Software\BJ\BJ's How to Get...\BJ's Basic Calender", "Calender Left") = "Error" Then
Me.Left = 4260
frmColours.Left = Me.Left + 1200
Else
Me.Left = GetBinaryValue("HKEY_CURRENT_USER\Software\BJ\BJ's How to Get...\BJ's Basic Calender", "Calender Left")
frmColours.Left = Me.Left + 1200
End If

If GetBinaryValue("HKEY_CURRENT_USER\Software\BJ\BJ's How to Get...\BJ's Basic Calender", "Calender Top") = "Error" Then
Me.Top = 15
Else
Me.Top = GetBinaryValue("HKEY_CURRENT_USER\Software\BJ\BJ's How to Get...\BJ's Basic Calender", "Calender Top")
End If

If GetBinaryValue("HKEY_CURRENT_USER\Software\BJ\BJ's How to Get...\BJ's Basic Calender", "Calender Height") = "Error" Then
Me.Height = 6030
Else
Me.Height = GetBinaryValue("HKEY_CURRENT_USER\Software\BJ\BJ's How to Get...\BJ's Basic Calender", "Calender Height")
End If

If GetBinaryValue("HKEY_CURRENT_USER\Software\BJ\BJ's How to Get...\BJ's Basic Calender", "Windows Time Top") = "Error" Then
Me.Frame4.Top = 4230
Else
Me.Frame4.Top = GetBinaryValue("HKEY_CURRENT_USER\Software\BJ\BJ's How to Get...\BJ's Basic Calender", "Windows Time Top")
End If

If GetBinaryValue("HKEY_CURRENT_USER\Software\BJ\BJ's How to Get...\BJ's Basic Calender", "Show\Hide TimeZone") = "Error" Then
Me.mnuShowHideTimeZone.Caption = "Hide TimeZone"
Else
Me.mnuShowHideTimeZone.Caption = GetBinaryValue("HKEY_CURRENT_USER\Software\BJ\BJ's How to Get...\BJ's Basic Calender", "Show\Hide TimeZone")
End If

'If Me.Height = 3840 Then
'mnuShowHideTimeZone.Enabled = False
'mnuShowHideWindowsTime.Enabled = False
'Else
'If Me.Height = 4590 Then
'mnuShowHideTimeZone.Enabled = False
'Else
'Me.Height = 5175
'mnuShowHideWindowsTime.Enabled = False
'End If
'End If

If GetStringValue("HKEY_CURRENT_USER\Software\BJ\BJ's How to Get...\BJ's Basic Calender", "Calender Text Colour") = "Error" Then
GoTo errors:
errors:
    Me.BackColor = frmColours.BackColor
    Me.Frame1.BackColor = frmColours.BackColor
    Me.Frame2.BackColor = frmColours.BackColor
    Me.Frame3.BackColor = frmColours.BackColor
    Me.Frame4.BackColor = frmColours.BackColor
    Me.lblMonth.ForeColor = vbBlack
    Me.lblYear.ForeColor = vbBlack
    Me.lblTime.ForeColor = vbBlack
    Me.lblDayofYear.ForeColor = vbBlack
    Me.lblWeekofYear.ForeColor = vbBlack
    Me.Label1.ForeColor = vbBlack
    Me.Label2.ForeColor = vbBlack
    Me.Label3.ForeColor = &H80FF&
    Me.Label4.ForeColor = vbBlack
    Me.Label6.ForeColor = vbBlack
    Me.Label7.ForeColor = vbBlack
    Me.Label8.ForeColor = vbBlack
    Me.Label9.ForeColor = &H80FF&
    Me.lblHrs.ForeColor = &HFFFF&
    Me.lblMins.ForeColor = &HFFFF&
    Me.lblDays.ForeColor = &HFFFF&
    Me.lblSecs.ForeColor = &HFFFF&
    frmColours.CurrentTextColor.ForeColor = Me.Label1.ForeColor
    frmColours.CurrentBackColor.BackColor = Me.Frame2.BackColor
    frmColours.CurrentTextColor.BackColor = Me.Frame2.BackColor
    frmColours.CurrentBackColor.ForeColor = Me.Label1.ForeColor
SetStringValue "HKEY_CURRENT_USER\Software\BJ\BJ's How to Get...\BJ's Basic Calender", "Calender Text Colour", Me.lblMonth.ForeColor
SetStringValue "HKEY_CURRENT_USER\Software\BJ\BJ's How to Get...\BJ's Basic Calender", "Calender Back Colour", Me.BackColor
    Else
    
    Me.lblMonth.ForeColor = GetStringValue("HKEY_CURRENT_USER\Software\BJ\BJ's How to Get...\BJ's Basic Calender", "Calender Text Colour")
    Me.lblYear.ForeColor = GetStringValue("HKEY_CURRENT_USER\Software\BJ\BJ's How to Get...\BJ's Basic Calender", "Calender Text Colour")
    Me.lblTime.ForeColor = GetStringValue("HKEY_CURRENT_USER\Software\BJ\BJ's How to Get...\BJ's Basic Calender", "Calender Text Colour")
    Me.lblDayofYear.ForeColor = GetStringValue("HKEY_CURRENT_USER\Software\BJ\BJ's How to Get...\BJ's Basic Calender", "Calender Text Colour")
    Me.lblWeekofYear.ForeColor = GetStringValue("HKEY_CURRENT_USER\Software\BJ\BJ's How to Get...\BJ's Basic Calender", "Calender Text Colour")
    Me.Label1.ForeColor = GetStringValue("HKEY_CURRENT_USER\Software\BJ\BJ's How to Get...\BJ's Basic Calender", "Calender Text Colour")
    Me.Label2.ForeColor = GetStringValue("HKEY_CURRENT_USER\Software\BJ\BJ's How to Get...\BJ's Basic Calender", "Calender Text Colour")
    Me.Label3.ForeColor = &H80FF&
    Me.Label4.ForeColor = GetStringValue("HKEY_CURRENT_USER\Software\BJ\BJ's How to Get...\BJ's Basic Calender", "Calender Text Colour")
    Me.Label6.ForeColor = GetStringValue("HKEY_CURRENT_USER\Software\BJ\BJ's How to Get...\BJ's Basic Calender", "Calender Text Colour")
    Me.Label7.ForeColor = GetStringValue("HKEY_CURRENT_USER\Software\BJ\BJ's How to Get...\BJ's Basic Calender", "Calender Text Colour")
    Me.Label8.ForeColor = GetStringValue("HKEY_CURRENT_USER\Software\BJ\BJ's How to Get...\BJ's Basic Calender", "Calender Text Colour")
    Me.Label9.ForeColor = &H80FF&
    Me.lblHrs.ForeColor = &HFFFF&
    Me.lblMins.ForeColor = &HFFFF&
    Me.lblDays.ForeColor = &HFFFF&
    Me.lblSecs.ForeColor = &HFFFF&
Me.BackColor = GetStringValue("HKEY_CURRENT_USER\Software\BJ\BJ's How to Get...\BJ's Basic Calender", "Calender Back Colour")
Frame1.BackColor = Me.BackColor
Frame2.BackColor = Frame1.BackColor
Frame3.BackColor = Frame2.BackColor
Frame4.BackColor = Frame3.BackColor
    frmColours.CurrentTextColor.ForeColor = Me.Label1.ForeColor
    frmColours.CurrentBackColor.BackColor = Me.Frame2.BackColor
    frmColours.CurrentTextColor.BackColor = Me.Frame2.BackColor
    frmColours.CurrentBackColor.ForeColor = Me.Label1.ForeColor
SetStringValue "HKEY_CURRENT_USER\Software\BJ\BJ's How to Get...\BJ's Basic Calender", "Calender Text Colour", Me.lblMonth.ForeColor
SetStringValue "HKEY_CURRENT_USER\Software\BJ\BJ's How to Get...\BJ's Basic Calender", "Calender Back Colour", Me.BackColor
  
    End If

Call bjsTimeZonesWin ' Gets your TimeZone Info...

mnuHour.Caption = GetStringValue("HKEY_CURRENT_USER\Software\BJ\BJ's How to Get...\BJ's Basic Calender", "Time Format")
If mnuHour.Caption = "Error" Then
mnuHour.Caption = "12 Hour Time"
End If
If mnuHour.Caption = "12 Hour Time" Then
MyTime = Format(Time, "hh:mm:ss ampm") '04:39:58 PM
mnuHour.Caption = "24 Hour Time"
Else
mnuHour.Caption = "24 Hour Time"
MyTime = Format(Time, "HH:MM:ss") '16:39:58
mnuHour.Caption = "12 Hour Time"
End If

'get current res and change screen resolution
'remove the 3 "'''" for the one you wamt to use
'''screenx = Screen.Width / Screen.TwipsPerPixelX
'''screeny = Screen.Height / Screen.TwipsPerPixelY
'used to hide the taskbar
'hWnd1 = FindWindow("Shell_traywnd", "")
'Call SetWindowPos(hWnd1, 0, 0, 0, 0, 0, SWP_HIDEWINDOW)
'change res if not right

'''Dim Msg, Style, Title, Response, MyString

'remove ''' if you want screen res at 640 X 480
'''Msg = "Do you want BJ's Basic Calender" & vbNewLine & _
"to change your screen resolution from: " & vbNewLine & vbNewLine & _
screenx & " X " & screeny & vbNewLine & _
"       to:" & vbNewLine & _
"640 X 480 ?"   ' Define message.

'remove ''' if you want screen res at 800 X 600
'''Msg = "Do you want BJ's Basic Calender" & vbNewLine & _
"to change your screen resolution from: " & vbNewLine & vbNewLine & _
screenx & " X " & screeny & vbNewLine & _
"       to:" & vbNewLine & _
"800 X 600 ?"   ' Define message.

'remove ''' if you want screen res at 1024 X 768
'''Msg = "Do you want BJ's Basic Calender" & vbNewLine & _
"to change your screen resolution from: " & vbNewLine & vbNewLine & _
screenx & " X " & screeny & vbNewLine & _
"       to:" & vbNewLine & _
"1024 X 768 ?"   ' Define message.

'''Style = vbYesNo + vbInformation + vbDefaultButton1   ' Define buttons.
'''Title = "Change Screen Resolution"   ' Define title.
      ' Display message.
'''Response = MsgBox(Msg, Style, Title)
'''If Response = vbYes Then   ' User chose Yes.
'''   MyString = "Yes"   ' Perform some action.

'remove ''' if you want screen res at 640 X 480
'''If screenx <> 640 Then Call ChangeRes(640, 480)

'remove ''' if you want screen res at 800 X 600
'''If screenx <> 800 Then Call ChangeRes(800, 600)

'remove ''' if you want screen res at 1024 X 768
'''If screenx <> 1024 Then Call ChangeRes(1024, 768)

'''Else   ' User chose No.
'''   MyString = "No"   ' Perform some action.
'''End If

    mhBaseCursor = GetClassLong((hwnd), GCL_HCURSOR)

End Sub

Private Sub Form_Unload(Cancel As Integer)
If screenx <> (Screen.Width / Screen.TwipsPerPixelX) Then Call ChangeRes((screenx), (screeny))
Call SetWindowPos(hWnd1, 0, 0, 0, 0, 0, SWP_SHOWWINDOW)
CleanUp

End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
    If Button And vbRightButton _
        Then PopupMenu mnuPopup

    
'    If Button And vbMiddleButton _
'        Then frmColours.Show
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)

Static Message As Long
Static RR As Boolean
    
    'x is the current mouse location along the x-axis
    Message = X / Screen.TwipsPerPixelX
    
    If RR = False Then
        RR = True
        Select Case Message
            ' Left double click (This should bring up a dialog box)
            Case WM_LBUTTONDBLCLK
    Call Shell_NotifyIcon(NIM_DELETE, TrayIcon)
                Me.Show
                mnuHide.Caption = "Hide"

            ' Right button up (This should bring up a menu)
            Case WM_RBUTTONUP
                Me.PopupMenu mnuPopup
        End Select
        RR = False
    End If
End Sub

Private Sub Frame1_DblClick()
   Call ControlPanels("rundll32.exe shell32.dll,Control_RunDLL timedate.cpl")
End Sub

Private Sub Frame1_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
    If Button And vbRightButton _
        Then PopupMenu mnuPopup
    
'    If Button And vbMiddleButton _
'        Then frmColours.Show
End Sub

Private Sub Frame2_DblClick()
   Call ControlPanels("rundll32.exe shell32.dll,Control_RunDLL timedate.cpl")
End Sub

Private Sub Frame2_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
    If Button And vbRightButton _
        Then PopupMenu mnuPopup
    
'    If Button And vbMiddleButton _
'        Then frmColours.Show
End Sub

Private Sub Frame3_DblClick()
   Call ControlPanels("rundll32.exe shell32.dll,Control_RunDLL timedate.cpl,,1")
End Sub

Private Sub Frame3_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
    If Button And vbRightButton _
        Then PopupMenu mnuPopup
    
'    If Button And vbMiddleButton _
'        Then frmColours.Show
End Sub

Private Sub Frame4_DblClick()
   Call ControlPanels("rundll32.exe shell32.dll,Control_RunDLL timedate.cpl,,1")
   
End Sub

Private Sub Frame4_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
    If Button And vbRightButton _
        Then PopupMenu mnuPopup
    
'    If Button And vbMiddleButton _
'        Then frmColours.Show
End Sub


Private Sub Label1_DblClick()
   Call ControlPanels("rundll32.exe shell32.dll,Control_RunDLL timedate.cpl")
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
    If Button And vbRightButton _
        Then PopupMenu mnuPopup
    
'    If Button And vbMiddleButton _
'        Then frmColours.Show
End Sub

Private Sub Label2_DblClick()
   Call ControlPanels("rundll32.exe shell32.dll,Control_RunDLL timedate.cpl")
End Sub

Private Sub Label2_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
    If Button And vbRightButton _
        Then PopupMenu mnuPopup
    
'    If Button And vbMiddleButton _
'        Then frmColours.Show
End Sub

Private Sub Label3_DblClick()
   Call ControlPanels("rundll32.exe shell32.dll,Control_RunDLL timedate.cpl")
End Sub

Private Sub Label9_DblClick()
   Call ControlPanels("rundll32.exe shell32.dll,Control_RunDLL timedate.cpl")
End Sub

Private Sub Label3_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
    If Button And vbRightButton _
        Then PopupMenu mnuPopup
    
'    If Button And vbMiddleButton _
'        Then frmColours.Show
    
End Sub

Private Sub Label9_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
    If Button And vbRightButton _
        Then PopupMenu mnuPopup
    
'    If Button And vbMiddleButton _
'        Then frmColours.Show
    
End Sub

Private Sub Label4_DblClick()
   Call ControlPanels("rundll32.exe shell32.dll,Control_RunDLL timedate.cpl,,1")
End Sub

Private Sub Label4_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
    If Button And vbRightButton _
        Then PopupMenu mnuPopup

    
'    If Button And vbMiddleButton _
'        Then frmColours.Show
End Sub

Private Sub Label6_DblClick()
   
   Call ControlPanels("rundll32.exe shell32.dll,Control_RunDLL timedate.cpl")
End Sub

Private Sub Label6_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
    If Button And vbRightButton _
        Then PopupMenu mnuPopup


    
'    If Button And vbMiddleButton _
'        Then frmColours.Show
End Sub

Private Sub Label7_DblClick()
   Call ControlPanels("rundll32.exe shell32.dll,Control_RunDLL timedate.cpl")
End Sub

Private Sub Label7_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
    If Button And vbRightButton _
        Then PopupMenu mnuPopup
    
'    If Button And vbMiddleButton _
'        Then frmColours.Show

End Sub

Private Sub Label8_DblClick()
   Call ControlPanels("rundll32.exe shell32.dll,Control_RunDLL timedate.cpl")
End Sub

Private Sub Label8_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
    If Button And vbRightButton _
        Then PopupMenu mnuPopup
    
'    If Button And vbMiddleButton _
'        Then frmColours.Show

End Sub

Private Sub lblDay_DblClick()
   Call ControlPanels("rundll32.exe shell32.dll,Control_RunDLL timedate.cpl")
End Sub

Private Sub lblDay_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
    If Button And vbRightButton _
        Then PopupMenu mnuPopup
    
'    If Button And vbMiddleButton _
'        Then frmColours.Show

End Sub

Private Sub lblDayofYear_DblClick()
   Call ControlPanels("rundll32.exe shell32.dll,Control_RunDLL timedate.cpl")
End Sub

Private Sub lblDayofYear_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
    If Button And vbRightButton _
        Then PopupMenu mnuPopup
    
'    If Button And vbMiddleButton _
'        Then frmColours.Show
End Sub

Private Sub lblDays_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
    If Button And vbRightButton _
        Then PopupMenu mnuPopup
    
'    If Button And vbMiddleButton _
'        Then frmColours.Show
End Sub

Private Sub Picture1_DblClick()
   Call ControlPanels("rundll32.exe shell32.dll,Control_RunDLL timedate.cpl")
End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
    If Button And vbRightButton _
        Then PopupMenu mnuPopup
    
'    If Button And vbMiddleButton _
'        Then frmColours.Show
End Sub


Public Function DoAtomicTime() As Boolean
'*****************************************************
'Purpose: Synchronizes System Clock with US Naval Atomic Clock

'Returns: True if Successful, false otherwise. Also saves time to
'          Global Variable m_StrAtomicTime

Dim TZI As TIME_ZONE_INFORMATION
Dim X As Single
Dim Ntime As Date
Dim i, n As Integer
Dim lRet As Long
Dim sRet As String
Dim bDaylightSavings As Boolean
Dim iMonthNow As Integer, iDayNow As Integer
Dim iStandardMonth As Integer, iStandardDay As Integer
Dim iDaylightMonth As Integer, iDaylightDay As Integer

m_StrAtomicTime = ""
Call GetTimeZoneInformation(TZI)


'Determine if we need to account for daylight savings time.
'this is precise to the day, not to the hour, so if you need
'more precision, change this code to consider wHour, wMinute
'and even wSecond Parameter
iStandardMonth = TZI.StandardDate.wMonth
iStandardDay = TZI.StandardDate.wDay

iDaylightMonth = TZI.DaylightDate.wMonth
iDaylightDay = TZI.DaylightDate.wDay

iMonthNow = Month(Now)
iDayNow = Day(Now)

If iStandardMonth = iMonthNow Then
    bDaylightSavings = (iDayNow < iStandardDay)
ElseIf iDaylightMonth = iMonthNow Then
    bDaylightSavings = (iDayNow >= iDaylightDay)
Else
    If iDaylightMonth < iStandardMonth Then
        bDaylightSavings = iMonthNow > iDaylightMonth _
            And iMonthNow < iStandardMonth
    Else
        bDaylightSavings = iMonthNow > iDaylightMonth _
           Or iMonthNow < iStandardMonth
    End If
End If
lRet = TZI.Bias

If bDaylightSavings Then
    lRet = lRet + TZI.DaylightBias
Else
    lRet = lRet + TZI.StandardBias
End If

X = lRet / 60
X = (X / 24)

Dim Msg, Style, Title, Response, MyString
Msg = "Do you want to Update your System Time with the Atomic Time?"   ' Define message.
Style = vbYesNo + vbInformation + vbDefaultButton1  ' Define buttons.
Title = "System Time Update"   ' Define title.
      ' Display message.
Response = MsgBox(Msg, Style, Title)
If Response = vbNo Then   ' User chose Yes.
   MyString = "No"   ' Perform some action.
DoAtomicTime = False
Exit Function
'-------------------------------------
Else   ' User chose No.
   MyString = "Yes"   ' Perform some action.
'-------------------------------------
'Else
'It will only go this far if you are connected to the internet
sRet = OpenURL("http://tycho.usno.navy.mil/cgi-bin/timer.pl")
i = InStr(1, sRet, "Universal")
If i <> 0 Then
 sRet = Left$(sRet, i) 'shorten the html string
 n = InStrRev(sRet, ",")
  If n <> 0 Then
   Ntime = CDate(Trim(Mid$(sRet, (n + 1), (i - (n + 1)))))
MsgBox "Atempting Synchronization with your time and the Atomic Clock...", vbInformation, "System Time update - Synchronizing" 'Update the status label
   Time = (Ntime - X)
   m_StrAtomicTime = CStr(Time)
   DoAtomicTime = True
  End If
End If
End If



End Function

Private Function OpenURL(ByVal sUrl As String) As String
'****************************************************
'From http://www.freevbcode.com/ShowCode.Asp?ID=1252
'*****************************************************

    Dim hOpen               As Long
    Dim hOpenUrl            As Long
    Dim bDoLoop             As Boolean
    Dim bRet                As Boolean
    Dim sReadBuffer         As String * 2048
    Dim lNumberOfBytesRead  As Long
    Dim sBuffer             As String

hOpen = InternetOpen(scUserAgent, INTERNET_OPEN_TYPE_PRECONFIG, _
    vbNullString, vbNullString, 0)

hOpenUrl = InternetOpenUrl(hOpen, sUrl, vbNullString, 0, _
   INTERNET_FLAG_RELOAD, 0)

    bDoLoop = True
    While bDoLoop
        sReadBuffer = vbNullString
        bRet = InternetReadFile(hOpenUrl, sReadBuffer, _
           Len(sReadBuffer), lNumberOfBytesRead)
        sBuffer = sBuffer & Left$(sReadBuffer, _
             lNumberOfBytesRead)
        If Not CBool(lNumberOfBytesRead) Then bDoLoop = False
    Wend

    If hOpenUrl <> 0 Then InternetCloseHandle (hOpenUrl)
    If hOpen <> 0 Then InternetCloseHandle (hOpen)
    OpenURL = sBuffer

End Function

Private Sub lblDisClock_Click()
Dim bSetTime As String, bAtomicTime As String

SetStringValue "HKEY_CURRENT_USER\Software\BJ\BJ's How to Get...\BJ's Basic Calender", "Current Time", Time
 m_StrAtomicTime = GetStringValue("HKEY_CURRENT_USER\Software\BJ\BJ's How to Get...\BJ's Basic Calender", "Atomic Time")

If DoAtomicTime Then
 bSetTime = GetStringValue("HKEY_CURRENT_USER\Software\BJ\BJ's How to Get...\BJ's Basic Calender", "Current Time")
'StartCursorAnimation
MsgBox "Your system time has been reset:" & vbNewLine & vbNewLine & _
"Your System Time Now is: - " & m_StrAtomicTime & vbNewLine & vbNewLine & _
"Your System Time was: - " & bSetTime, vbInformation, "System Time update - Synchronized"

'StopCursorAnimation
Else
Dim Msg, Style, Title, Response, MyString
Msg = "The attempt to synchronize your system time has failed."   ' Define message.
Style = vbOKOnly + vbCritical + vbDefaultButton1  ' Define buttons.
Title = "System Time update - Failed"   ' Define title.
      ' Display message.
Response = MsgBox(Msg, Style, Title)
If Response = vbOKOnly Then   ' User chose Yes.
   MyString = "No"   ' Perform some action.
'StopCursorAnimation
Exit Sub
'-------------------------------------
End If
End If
End Sub

Private Sub lblHrs_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
    If Button And vbRightButton _
        Then PopupMenu mnuPopup


    
'    If Button And vbMiddleButton _
'        Then frmColours.Show
End Sub

Private Sub lblMins_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
    If Button And vbRightButton _
        Then PopupMenu mnuPopup


    
'    If Button And vbMiddleButton _
'        Then frmColours.Show
End Sub

Private Sub lblMonth_DblClick()
   Call ControlPanels("rundll32.exe shell32.dll,Control_RunDLL timedate.cpl")
End Sub

Private Sub lblMonth_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
    If Button And vbRightButton _
        Then PopupMenu mnuPopup
    
'    If Button And vbMiddleButton _
'        Then frmColours.Show
End Sub

Private Sub lblNumber_DblClick()
   Call ControlPanels("rundll32.exe shell32.dll,Control_RunDLL timedate.cpl")
End Sub

Private Sub lblNumber_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
    If Button And vbRightButton _
        Then PopupMenu mnuPopup
    
'    If Button And vbMiddleButton _
'        Then frmColours.Show
End Sub

Private Sub lblSecs_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
    If Button And vbRightButton _
        Then PopupMenu mnuPopup
    
'    If Button And vbMiddleButton _
'        Then frmColours.Show
End Sub

Private Sub lblTime_DblClick()
   Call ControlPanels("rundll32.exe shell32.dll,Control_RunDLL timedate.cpl")
End Sub

Private Sub lblTime_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
    If Button And vbRightButton _
        Then PopupMenu mnuPopup
    
'    If Button And vbMiddleButton _
'        Then frmColours.Show
End Sub

Private Sub lblWeekofYear_DblClick()
   Call ControlPanels("rundll32.exe shell32.dll,Control_RunDLL timedate.cpl")
End Sub

Private Sub lblWeekofYear_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
    If Button And vbRightButton _
        Then PopupMenu mnuPopup
    
'    If Button And vbMiddleButton _
'        Then frmColours.Show
End Sub

Private Sub lblYear_DblClick()
   Call ControlPanels("rundll32.exe shell32.dll,Control_RunDLL timedate.cpl")
End Sub

Private Sub lblYear_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
    If Button And vbRightButton _
        Then PopupMenu mnuPopup
    
'    If Button And vbMiddleButton _
'        Then frmColours.Show
End Sub

Private Sub mnuAtomicTime_Click()
lblDisClock_Click
End Sub

Private Sub mnuColor_Click()
If Me.mnuHide.Caption = "Show" Then
mnuHide_Click
frmColours.Left = frmBasicCalender.Left + 1200
frmColours.Show
Else
frmColours.Show
End If
End Sub


Private Sub mnuHide_Click()
If mnuHide.Caption = "Show" Then
Me.Show
mnuHide.Caption = "Hide"
Else
If mnuHide.Caption = "Hide" Then
Me.Hide
mnuHide.Caption = "Show"
End If
End If

If frmColours.Visible = True Then
Unload frmColours
End If
Call tIcon
End Sub

Private Sub mnuHour_Click()

mnuHour.Caption = GetStringValue("HKEY_CURRENT_USER\Software\BJ\BJ's How to Get...\BJ's Basic Calender", "Time Format")
If mnuHour.Caption = "Error" Then
mnuHour.Caption = "12 Hour Time"
End If
If mnuHour.Caption = "12 Hour Time" Then
MyTime = Format(Time, "hh:mm:ss ampm") '04:39:58 PM
mnuHour.Caption = "24 Hour Time"
Else
If mnuHour.Caption = "24 Hour Time" Then
MyTime = Format(Time, "HH:MM:ss") '16:39:58
mnuHour.Caption = "12 Hour Time"
End If
End If
lblTime = MyTime
lblTime.Refresh
SetStringValue "HKEY_CURRENT_USER\Software\BJ\BJ's How to Get...\BJ's Basic Calender", "Time Format", mnuHour.Caption

End Sub

Private Sub AboutBox(hwnd As Long)
'-----------------------------------------------------------------
    ' Show help about dialog...
    ShellAbout hwnd, App.EXEName, _
               vbCrLf & "BJ. E-Mail me: bryce3@bigpond.com", frmBasicCalender.Icon
'-----------------------------------------------------------------
End Sub

Private Sub mnuAbout_Click()
   AboutBox Me.hwnd
End Sub

Private Sub mnuDateTime_Click()
   Call ControlPanels("rundll32.exe shell32.dll,Control_RunDLL timedate.cpl")
End Sub

Private Sub mnuEMail_Click()
'Info to E-Mail someone
' Change to what you want.   mailto: can be changed to www.
    ShellExecute 0, "Open", "mailto:bryce3@bigpond.com?subject=Info about BJ's How to Get... Basic Calendar.&body=Type here the message that you want to send me.", "", "", vbMaximizedFocus

End Sub

Private Sub mnuExit_Click()
CleanUp
End Sub

Private Sub mnuRunMenu_Click()
If mnuRunMenu.Caption = "Add to Run Command" Then
CreateKey "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\BJ'sBC.exe"
SetStringValue "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\BJ'sBC.exe", "", App.Path & "\" & App.EXEName & ".exe"
MsgBox "You can now start BJ's Basic Calender" & vbNewLine & "from your Run Command by typing in" & vbNewLine & vbNewLine & "BJ'sBC", vbInformation, App.EXEName & " is now in your Run Command"
mnuRunMenu.Caption = "Remove from Run Command"

ElseIf mnuRunMenu.Caption = "Remove from Run Command" Then
DeleteKey "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths", "BJ'sBC.exe"
MsgBox "BJ'sBC has been removed from your Run Command." & vbNewLine & "If you type in BJ'sBC it will not work.", vbInformation, App.EXEName & " has been removed."
mnuRunMenu.Caption = "Add to Run Command"
End If
SetStringValue "HKEY_CURRENT_USER\Software\BJ\BJ's How to Get...\BJ's Basic Calender", "Run Command", mnuRunMenu.Caption
End Sub

Private Sub mnuShowHideTimeZone_Click()

If Me.mnuShowHideTimeZone.Caption = GetStringValue("HKEY_CURRENT_USER\Software\BJ\BJ's How to Get...\BJ's Basic Calender", "Show\Hide TimeZone") = "Error" Then
Me.mnuShowHideTimeZone.Caption = "Hide TimeZone"
End If

If Me.mnuShowHideTimeZone.Caption = "Show TimeZone" Then
Me.mnuShowHideTimeZone.Caption = "Hide TimeZone"
Me.Frame4.Top = 4160
Me.Height = 6030
SetStringValue "HKEY_CURRENT_USER\Software\BJ\BJ's How to Get...\BJ's Basic Calender", "Windows Time Top", Me.Frame4.Top
SetStringValue "HKEY_CURRENT_USER\Software\BJ\BJ's How to Get...\BJ's Basic Calender", "Calender Height", Me.Height
SetStringValue "HKEY_CURRENT_USER\Software\BJ\BJ's How to Get...\BJ's Basic Calender", "Show\Hide TimeZone", "Show TimeZone"

Else

If Me.mnuShowHideTimeZone.Caption = "Hide TimeZone" Then
Me.mnuShowHideTimeZone.Caption = "Show TimeZone"
Me.Frame4.Top = 3420
Me.Height = 5295
SetStringValue "HKEY_CURRENT_USER\Software\BJ\BJ's How to Get...\BJ's Basic Calender", "Windows Time Top", Me.Frame4.Top
SetStringValue "HKEY_CURRENT_USER\Software\BJ\BJ's How to Get...\BJ's Basic Calender", "Calender Height", Me.Height
SetStringValue "HKEY_CURRENT_USER\Software\BJ\BJ's How to Get...\BJ's Basic Calender", "Show\Hide TimeZone", "Hide TimeZone"
End If
End If

End Sub

Private Sub mnuTimeZone_Click()
Label4_DblClick
End Sub

Private Sub timDisplay_Timer()
SetStringValue "HKEY_CURRENT_USER\Software\BJ\BJ's How to Get...\BJ's Basic Calender", "Calender Text Colour", Me.lblMonth.ForeColor
SetStringValue "HKEY_CURRENT_USER\Software\BJ\BJ's How to Get...\BJ's Basic Calender", "Calender Back Colour", Me.BackColor


'Checks for leap year
Dim DayinYear
DayinYear = Format(Today, "yyyy")
 If IsLeapYear(DayinYear) = True Then
DayinYear = 366
Else
DayinYear = 365
End If



mnuHour.Caption = GetStringValue("HKEY_CURRENT_USER\Software\BJ\BJ's How to Get...\BJ's Basic Calender", "Time Format")
If mnuHour.Caption = "12 Hour Time" Then
MyTime = Format(Time, "hh:mm:ss ampm") '04:39:58 PM
mnuHour.Caption = "24 Hour Time"
 Else
mnuHour.Caption = "24 Hour Time"
MyTime = Format(Time, "HH:MM:ss") '16:39:58
mnuHour.Caption = "12 Hour Time"
End If


'Current time
Today = Now
'Change format of day to eg... Sunday
lblDay.Caption = Format(Today, "dddd")
'Change format of month to eg... September
lblMonth.Caption = Format(Today, "mmmm")
'Change format of year to eg... 2002
lblYear.Caption = Format(Today, "yyyy")
'Change format of date to eg... 02
lblNumber.Caption = Format(Today, "dd")
'Change format of time to eg... 22:34:45 PM or 10:34:45 PM
lblTime.Caption = MyTime
'Change format of day of year to eg... 234 of 365
lblDayofYear.Caption = Format(Today, "y" & " of " & DayinYear)
'Change format of week of year to eg... 22 of 52
lblWeekofYear.Caption = Format(Today, "ww" & " of 52")

'Calculates remaining days and weeks in year
If DayinYear = 366 Then
Label7.Caption = "Days left:             " & DayinYear - Format(Today, "y")
Label8.Caption = "Weeks left:             " & 52 - Format(Today, "ww")

ElseIf DayinYear = 365 Then
Label7.Caption = "Days left:             " & DayinYear - Format(Today, "y")
Label8.Caption = "Weeks left:             " & 52 - Format(Today, "ww")
End If

'monday to friday will be shown in blue
If lblDay.Caption = "Monday" Then
lblDay.ForeColor = &HFF0000
lblNumber.ForeColor = &HC00000
Label6.ForeColor = &HFF0000
ElseIf lblDay.Caption = "Tuesday" Then
lblDay.ForeColor = &HFF0000
lblNumber.ForeColor = &HC00000
Label6.ForeColor = &HFF0000
ElseIf lblDay.Caption = "Wednesday" Then
lblDay.ForeColor = &HFF0000
lblNumber.ForeColor = &HC00000
Label6.ForeColor = &HFF0000
ElseIf lblDay.Caption = "Thursday" Then
lblDay.ForeColor = &HFF0000
lblNumber.ForeColor = &HC00000
Label6.ForeColor = &HFF0000
ElseIf lblDay.Caption = "Friday" Then
lblDay.ForeColor = &HFF0000
lblNumber.ForeColor = &HC00000
Label6.ForeColor = &HFF0000
'saturday will be shown in green
ElseIf lblDay.Caption = "Saturday" Then
lblDay.ForeColor = &HFF00&
lblNumber.ForeColor = &HC000&
Label6.ForeColor = &HC000&
'sunday will be shown in red
ElseIf lblDay.Caption = "Sunday" Then
lblDay.ForeColor = &HFF&
lblNumber.ForeColor = &HC0&
Label6.ForeColor = &HC0&
End If
 Horoscopes

End Sub


Private Sub Timer1_Timer()
SetStringValue "HKEY_CURRENT_USER\Software\BJ\BJ's How to Get...\BJ's Basic Calender", "Calender Text Colour", Me.lblMonth.ForeColor
SetStringValue "HKEY_CURRENT_USER\Software\BJ\BJ's How to Get...\BJ's Basic Calender", "Calender Back Colour", Me.BackColor
'The following codes will rotate the image of the clock
RotateMyClock = RotateMyClock + 1
If RotateMyClock = 12 Then RotateMyClock = 0 'reset the clock
lblDisClock.Caption = lblPic(RotateMyClock).Caption
lblDisClock.ForeColor = RGB(Rnd * 255, Rnd * 255, Rnd * 255)
   Call bjsTimeZonesWin

End Sub



Private Sub tmrCheckTime_Timer()
SetStringValue "HKEY_CURRENT_USER\Software\BJ\BJ's How to Get...\BJ's Basic Calender", "Calender Text Colour", Me.lblMonth.ForeColor
SetStringValue "HKEY_CURRENT_USER\Software\BJ\BJ's How to Get...\BJ's Basic Calender", "Calender Back Colour", Me.BackColor
Dim Today As Variant
Dim MyDate, MyTime, MyDay, MyDay1, MyDay2, MyWeek, MyWeek1, MyWeek2
   Static AddBar As Integer, i As Integer
'Format to get the time windows has been open for
     
     ms = GetTickCount()
    days = ms \ MS_PER_DAY
    ms = ms - days * MS_PER_DAY
    hrs = ms \ MS_PER_HR
    ms = ms - hrs * MS_PER_HR
    mins = ms \ MS_PER_MIN
    ms = ms - mins * MS_PER_MIN
    secs = ms \ MS_PER_SEC
    ms = ms - secs * MS_PER_SEC
    
'The following will display for Eg... If you have 1 Second you will see
' 1 Seconds, with the format below it will show 1 Second, 2 Seconds etc...

If days = 1 Then
    lblDays.Caption = Format$(days) & " Day"
Else
lblDays.Caption = Format$(days) & " Days"
End If

If hrs = 1 Then
    lblHrs.Caption = Format$(hrs) & " Hour"
Else
    lblHrs.Caption = Format$(hrs) & " Hours"
End If

If mins = 1 Then
    lblMins.Caption = Format$(mins) & " Minute"
Else
lblMins.Caption = Format$(mins) & " Minutes"
End If
    
If secs = 1 Then
    lblSecs.Caption = Format$(secs) & " Second"
    Else
    lblSecs.Caption = Format$(secs) & " Seconds"
End If
    

'-------------------------------------------------------------------------------------------
End Sub

Private Sub mnuWinStart_Click()
'This section creates Shortcut (.lnk) in your Startup menu
If mnuWinStart.Caption = "Add to Startup" Then
mnuWinStart.Caption = "Remove from Startup"
Dim rc As Long                                      ' return code
    Dim sLnk As clsShellLink                               ' ShellLink class object
    Dim sfPath As String                                ' System folder path
    Dim Id As Long                                      ' ID of System folder...
    ' Create instance of Explorer's IShellLink Interface Base Class
    Set sLnk = New clsShellLink
    
    'Id = cmbSysFolders.ItemData(cmbSysFolders.ListIndex)  ' Get ID from combo box
    'To add .lnk current user change below to &H7
    If sLnk.GetSystemFolderPath(Me.hwnd, &H18, sfPath) Then ' Get system folder path from id
        SetDefaults sfPath                                ' Update UI with new path
    End If
    
    
    Set sLnk = New clsShellLink                            ' Create ShellLink Instance
    
    sLnk.CreateShellLink txtLinkName.Text & txtlnk.Text, txtExeName.Text, txtWorkDir.Text, txtCmdArgs.Text, txtLnkDesc.Text, txtIconFile.Text, CLng(txtIconIndex.Text), CLng(txtShowCmd.Text)        ' Create a ShellLink (ShortCut)
    
    Set sLnk = Nothing                                  ' Destroy object reference
SetStringValue "HKEY_CURRENT_USER\Software\BJ\BJ's How to Get...\BJ's Basic Calender", "Windows Startup", mnuWinStart.Caption
SetStringValue "HKEY_CURRENT_USER\Software\BJ\BJ's How to Get...\BJ's Basic Calender", "Windows Startup Path and File", txtLinkName.Text & txtlnk.Text
SetStringValue "HKEY_CURRENT_USER\Software\BJ\BJ's How to Get...\BJ's Basic Calender", ".exe Path", txtExeName.Text
    'MsgBox "The Shortcut " & vbNewLine & txtLinkName.Text & vbNewLine & " was created for you", vbInformation, App.EXEName & " - Shortcut"
    MsgBox "The Shortcut to BJ's Basic Calender.lnk was created for you in:" & vbNewLine & txtLinkName.Text, vbInformation, App.EXEName & " - Shortcut"
Else
'This section Deletes Shortcut (.lnk) to the Recycle Bin

If mnuWinStart.Caption = "Remove from Startup" Then
    mnuWinStart.Caption = "Add to Startup"
SetStringValue "HKEY_CURRENT_USER\Software\BJ\BJ's How to Get...\BJ's Basic Calender", "Windows Startup", mnuWinStart.Caption
Dim FileOperation As SHFILEOPSTRUCT
Dim lReturn As Long

    Set sLnk = New clsShellLink

With FileOperation

    If sLnk.GetSystemFolderPath(Me.hwnd, &H7, sfPath) Then ' Get system folder path from id
        SetDefaults sfPath                                ' Update UI with new path
'SetStringValue "HKEY_CURRENT_USER\Software\BJ\BJ's How to Get...\BJ's Basic Calender", "Your Startup Directory", sfPath

'If GetStringValue("HKEY_CURRENT_USER\Software\BJ\BJ's How to Get...\BJ's Basic Calender", "Your Startup Directory") = "Error" Then
    End If
    .wFunc = FO_DELETE
If FileExists(sfPath & "BJ's Basic Calender.lnk") = True Then
    .pFrom = sfPath & "BJ's Basic Calender.lnk"     'fichier sélectionné dans la liste
    .fFlags = FOF_ALLOWUNDO
    End If
End With
    Set sLnk = Nothing                                  ' Destroy object reference

    mnuWinStart.Caption = "Add to Startup"

lReturn = SHFileOperation(FileOperation)
End If
End If
End Sub

Private Sub SetDefaults(pth As String)
    
    Dim AppPath As String                                   ' Current Application path
    
    AppPath = App.Path                                      ' Get current path
    
    If (Right$(AppPath, 1) <> "\") Then AppPath = AppPath & "\" ' Fix application path if necessary
    If (Right$(pth, 1) <> "\") Then pth = pth & "\"         ' Fix path if necessary
    
'all the text boxes are hidden to the right of the form. just easier to do it this way
    txtlnk.Text = "BJ's Basic Calender.lnk"                 ' Set default Program group name
    txtLinkName.Text = pth                                  ' Create a full path name for link file
    txtExeName.Text = AppPath & App.EXEName & ".exe"        ' Create a full path name for applicaton exe name
    txtWorkDir.Text = AppPath                               ' Set default working directory
    txtLnkDesc.Text = "BJ's Basic Calender: Shows Current Date and Time and Windows Running Time." ' Set comment
    txtwHotKey.Text = "C"                                   ' Set hot key
    txtIconFile.Text = txtExeName.Text                      ' Set default IconFile name to default exename
    txtIconIndex.Text = CStr(1)                             '0=forms default icon. Set default Icon Index val
    txtShowCmd.Text = CStr(5)                               '3=Max 5=Normal 7=Min.  set default showcommand val

End Sub

Function FileExists(FileName As String) As Integer
    On Error Resume Next
    Dim X%
        X% = Len(Dir$(FileName))
    If Err Or X% = 0 Then
    FileExists = False
    mnuWinStart.Caption = "Add to Startup"
    Else
    FileExists = True
    mnuWinStart.Caption = "Remove from Startup"
    End If
End Function

Private Sub Horoscopes()
Dim Starsigns

If lblMonth = "March" And lblNumber > "20" Or lblMonth = "April" And lblNumber < "21" Then
Starsigns = "Aries"
Label9.Caption = "March 21 to April 20"
'Aries Dates: March 21 to April 20

ElseIf lblMonth = "April" And lblNumber > "20" Or lblMonth = "May" And lblNumber < "22" Then
Starsigns = "Taurus"
Label9.Caption = "April 21 to May 21"
'Taurus Dates: April 21 to May 21

ElseIf lblMonth = "May" And lblNumber > "21" Or lblMonth = "June" And lblNumber < "23" Then
Starsigns = "Gemini"
Label9.Caption = "May 22 to June 21"
'Gemini Dates: May 22 to June 21

ElseIf lblMonth = "June" And lblNumber > "22" Or lblMonth = "July" And lblNumber < "24" Then
Starsigns = "Cancer"
Label9.Caption = "June 22 to July 23"
'Cancer Dates: June 22 to July 23

ElseIf lblMonth = "July" And lblNumber > "23" Or lblMonth = "August" And lblNumber < "24" Then
Starsigns = "Leo"
Label9.Caption = "July 24 to August 23"
'Leo Dates: July 24 to August 23

ElseIf lblMonth = "August" And lblNumber > "23" Or lblMonth = "September" And lblNumber < "24" Then
Starsigns = "Virgo"
Label9.Caption = "August 24 to September 23"
'Virgo Dates: August 24 to September 23

ElseIf lblMonth = "September" And lblNumber > "23" Or lblMonth = "October" And lblNumber < "24" Then
Starsigns = "Libra"
Label9.Caption = "September 24 to October 23"
'Libra Dates: September 24 to October 23

ElseIf lblMonth = "October" And lblNumber > "23" Or lblMonth = "November" And lblNumber < "23" Then
Starsigns = "Scorpio"
Label9.Caption = "October 24 to November 22"
'Scorpio Dates: October 24 to November 22

ElseIf lblMonth = "November" And lblNumber > "22" Or lblMonth = "December" And lblNumber < "22" Then
Starsigns = "Sagittarius"
Label9.Caption = "November 23 to December 21"
'Sagittarius Dates: November 23 to December 21

ElseIf lblMonth = "December" And lblNumber > "21" Or lblMonth = "January" And lblNumber < "21" Then
Starsigns = "Capricorn"
Label9.Caption = "December 22 to January 20"
'Capricorn Dates: December 22 to January 20

ElseIf lblMonth = "January" And lblNumber > "20" Or lblMonth = "February" And lblNumber < "20" Then
Starsigns = "Aquarius"
Label9.Caption = "January 21 to February 19"
'Aquarius Dates: January 21 to February 19

ElseIf lblMonth = "February" And lblNumber > "19" Or lblMonth = "March" And lblNumber < "21" Then
Starsigns = "Pisces"
Label9.Caption = "February 20 to March 20"
'Pisces Dates: February 20 to March 20

End If

Label3 = Starsigns
Picture1.Picture = LoadPicture(App.Path & "\StarSign Pictures\" & Starsigns & ".jpg")

End Sub
Private Function bjsTimeZonesWin()
'all the timezone info... for Windows 95 types.
'This had to be done like this because in the Registry
' the timezone info... is like the next line.
'HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Control\TimeZoneInformation and to the right is "Afghanistan Standard Time"
'where in the timezone section it is like the next line.
'"HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Time Zones\Afghanistan" and to the right is "(GMT+04:30) Kabul"
'Windows 95 Versions show "Afghanistan" not "Afghanistan Standard Time" like in Windows NT Versions
'For Windows NT versions goto the end of this section

Dim bjTimeZoneTime

Label5 = GetStringValue("HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Control\TimeZoneInformation", "StandardName")
Me.Caption = "BJ's Basic Calender - " & Label5

If Label5 = "Afghanistan Standard Time" Then
Label4 = GetStringValue("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Time Zones\Afghanistan", "Display")

ElseIf Label5 = "Alaskan Standard Time" Then
Label4 = GetStringValue("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Time Zones\Alaskan", "Display")

ElseIf Label5 = "Arab Standard Time" Then
Label4 = GetStringValue("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Time Zones\Arab", "Display")

ElseIf Label5 = "Arabian Standard Time" Then
Label4 = GetStringValue("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Time Zones\Arabian", "Display")

ElseIf Label5 = "Atlantic Standard Time" Then
Label4 = GetStringValue("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Time Zones\Atlantic", "Display")

ElseIf Label5 = "AUS Central Standard Time" Then
Label4 = GetStringValue("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Time Zones\AUS Central", "Display")

ElseIf Label5 = "AUS Eastern Standard Time" Then
Label4 = GetStringValue("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Time Zones\AUS Eastern", "Display")

ElseIf Label5 = "Azores Standard Time" Then
Label4 = GetStringValue("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Time Zones\Azores", "Display")

ElseIf Label5 = "Canada Central  Standard Time" Then
Label4 = GetStringValue("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Time Zones\Canada Central", "Display")

ElseIf Label5 = "Caucasus Standard Time" Then
Label4 = GetStringValue("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Time Zones\Caucasus", "Display")

ElseIf Label5 = "Cen. Australia Standard Time" Then
Label4 = GetStringValue("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Time Zones\Cen. Australia", "Display")

ElseIf Label5 = "Central Standard Time" Then
Label4 = GetStringValue("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Time Zones\Central", "Display")

ElseIf Label5 = "Central Asia Standard Time" Then
Label4 = GetStringValue("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Time Zones\Central Asia", "Display")

ElseIf Label5 = "Central Europe Standard Time" Then
Label4 = GetStringValue("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Time Zones\Central Europe", "Display")

ElseIf Label5 = "Central European Standard Time" Then
Label4 = GetStringValue("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Time Zones\Central European", "Display")

ElseIf Label5 = "Central Pacific Standard Time" Then
Label4 = GetStringValue("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Time Zones\Central Pacific", "Display")

ElseIf Label5 = "China Standard Time" Then
Label4 = GetStringValue("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Time Zones\China", "Display")

ElseIf Label5 = "E. Africa Standard Time" Then
Label4 = GetStringValue("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Time Zones\E. Africa", "Display")

ElseIf Label5 = "E. Australia Standard Time" Then
Label4 = GetStringValue("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Time Zones\E. Australia", "Display")

ElseIf Label5 = "E. Europe Standard Time" Then
Label4 = GetStringValue("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Time Zones\E. Europe", "Display")

ElseIf Label5 = "E. South America Standard Time" Then
Label4 = GetStringValue("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Time Zones\E. South America", "Display")

ElseIf Label5 = "Eastern Standard Time" Then
Label4 = GetStringValue("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Time Zones\Eastern", "Display")

ElseIf Label5 = "Egypt Standard Time" Then
Label4 = GetStringValue("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Time Zones\Egypt", "Display")

ElseIf Label5 = "Ekaterinburg Standard Time" Then
Label4 = GetStringValue("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Time Zones\Ekaterinburg", "Display")

ElseIf Label5 = "Fiji Standard Time" Then
Label4 = GetStringValue("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Time Zones\Fiji", "Display")

ElseIf Label5 = "FLE Standard Time" Then
Label4 = GetStringValue("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Time Zones\FLE", "Display")

ElseIf Label5 = "GMT Standard Time" Then
Label4 = GetStringValue("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Time Zones\GMT", "Display")

ElseIf Label5 = "Greenwich Standard Time" Then
Label4 = GetStringValue("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Time Zones\Greenwich", "Display")

ElseIf Label5 = "GTB Standard Time" Then
Label4 = GetStringValue("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Time Zones\GTB", "Display")

ElseIf Label5 = "India Standard Time" Then
Label4 = GetStringValue("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Time Zones\India", "Display")

ElseIf Label5 = "Iran Standard Time" Then
Label4 = GetStringValue("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Time Zones\Iran", "Display")

ElseIf Label5 = "Jerusalem Standard Time" Then
Label4 = GetStringValue("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Time Zones\Israel", "Display")

ElseIf Label5 = "Korea Standard Time" Then
Label4 = GetStringValue("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Time Zones\Korea", "Display")

ElseIf Label5 = "Mexico Standard Time" Then
Label4 = GetStringValue("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Time Zones\Mexico", "Display")

ElseIf Label5 = "Mid-Atlantic Standard Time" Then
Label4 = GetStringValue("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Time Zones\Mid-Atlantic", "Display")

ElseIf Label5 = "Mountain Standard Time" Then
Label4 = GetStringValue("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Time Zones\Mountain", "Display")

ElseIf Label5 = "New Zealand Standard Time" Then
Label4 = GetStringValue("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Time Zones\New Zealand", "Display")

ElseIf Label5 = "Newfoundland Standard Time" Then
Label4 = GetStringValue("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Time Zones\Newfoundland", "Display")

ElseIf Label5 = "Pacific Standard Time" Then
Label4 = GetStringValue("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Time Zones\Pacific", "Display")

ElseIf Label5 = "Pacific SA Standard Time" Then
Label4 = GetStringValue("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Time Zones\Pacific SA", "Display")

ElseIf Label5 = "Romance Standard Time" Then
Label4 = GetStringValue("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Time Zones\Romance", "Display")

ElseIf Label5 = "Russian Standard Time" Then
Label4 = GetStringValue("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Time Zones\Russian", "Display")

ElseIf Label5 = "SA Eastern Standard Time" Then
Label4 = GetStringValue("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Time Zones\SA Eastern", "Display")

ElseIf Label5 = "SA Pacific Standard Time" Then
Label4 = GetStringValue("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Time Zones\SA Pacific", "Display")

ElseIf Label5 = "SA Western Standard Time" Then
Label4 = GetStringValue("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Time Zones\SA Western", "Display")

ElseIf Label5 = "SE Asia Standard Time" Then
Label4 = GetStringValue("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Time Zones\SE Asia", "Display")

ElseIf Label5 = "Singapore Standard Time" Then
Label4 = GetStringValue("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Time Zones\Singapore", "Display")

ElseIf Label5 = "South Africa Standard Time" Then
Label4 = GetStringValue("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Time Zones\South Africa", "Display")

ElseIf Label5 = "Sri Lanka Standard Time" Then
Label4 = GetStringValue("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Time Zones\Sri Lanka", "Display")

ElseIf Label5 = "Taipei Standard Time" Then
Label4 = GetStringValue("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Time Zones\Taipei", "Display")

ElseIf Label5 = "Tasmania Standard Time" Then
Label4 = GetStringValue("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Time Zones\Tasmania", "Display")

ElseIf Label5 = "Tokyo Standard Time" Then
Label4 = GetStringValue("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Time Zones\Tokyo", "Display")

ElseIf Label5 = "US Eastern Standard Time" Then
Label4 = GetStringValue("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Time Zones\US Eastern", "Display")

ElseIf Label5 = "US Mountain Standard Time" Then
Label4 = GetStringValue("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Time Zones\US Mountain", "Display")

ElseIf Label5 = "Vladivostok Standard Time" Then
Label4 = GetStringValue("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Time Zones\Vladivostok", "Display")

ElseIf Label5 = "W. Australia Standard Time" Then
Label4 = GetStringValue("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Time Zones\W. Australia", "Display")

ElseIf Label5 = "W. Europe Standard Time" Then
Label4 = GetStringValue("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Time Zones\W. Europe", "Display")

ElseIf Label5 = "West Asia Standard Time" Then
Label4 = GetStringValue("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Time Zones\West Asia", "Display")

ElseIf Label5 = "West Pacific Standard Time" Then
Label4 = GetStringValue("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Time Zones\West Pacific", "Display")

ElseIf Label5 = "Yakutsk Standard Time" Then
Label4 = GetStringValue("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Time Zones\Yakutsk", "Display")
End If
If Label4 = "Error" Then
GoTo bjsTimeZonesWinNT

Me.Refresh
'Windows NT section
bjsTimeZonesWinNT:
Label5 = GetStringValue("HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Control\TimeZoneInformation", "StandardName")

Label4.Caption = GetStringValue("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Time Zones\" & Label5, "Display")

Me.Caption = "BJ's Basic Calender - " & Label5
Me.Refresh
End If

SetStringValue "HKEY_CURRENT_USER\Software\BJ\BJ's How to Get...\BJ's Basic Calender", "Time Zone Info", Label5.Caption
SetStringValue "HKEY_CURRENT_USER\Software\BJ\BJ's How to Get...\BJ's Basic Calender", "Time Zone", Label4.Caption


End Function

