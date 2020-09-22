VERSION 5.00
Begin VB.Form frmIcon 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Tests the IShellLink Typelib Interface."
   ClientHeight    =   4335
   ClientLeft      =   420
   ClientTop       =   720
   ClientWidth     =   9195
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4335
   ScaleWidth      =   9195
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox pctIcon 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   540
      Left            =   360
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   36
      TabStop         =   0   'False
      Top             =   360
      Width           =   540
   End
   Begin VB.VScrollBar PicIconIndex 
      Height          =   495
      Left            =   960
      Max             =   100
      TabIndex        =   35
      Top             =   360
      Width           =   255
   End
   Begin VB.TextBox txtPath 
      Height          =   285
      Left            =   3120
      Locked          =   -1  'True
      TabIndex        =   32
      Top             =   600
      Width           =   5865
   End
   Begin VB.TextBox txtURLTarget 
      Height          =   285
      Left            =   3120
      TabIndex        =   29
      Top             =   7560
      Width           =   5895
   End
   Begin VB.TextBox txtURL 
      Height          =   285
      Left            =   3120
      TabIndex        =   28
      Top             =   7200
      Width           =   5895
   End
   Begin VB.CommandButton cmdURL 
      Caption         =   "Create .url"
      Height          =   375
      Left            =   120
      TabIndex        =   26
      Top             =   7200
      Width           =   1095
   End
   Begin VB.TextBox txtwHotKey 
      Height          =   285
      Left            =   3000
      TabIndex        =   24
      Top             =   6240
      Width           =   2295
   End
   Begin VB.TextBox txtLnkDesc 
      Height          =   285
      Left            =   3120
      TabIndex        =   23
      Top             =   2520
      Width           =   5895
   End
   Begin VB.TextBox txtProgramGroup 
      Height          =   285
      Left            =   3000
      TabIndex        =   20
      Top             =   6660
      Width           =   5865
   End
   Begin VB.CommandButton cmdCreateGroup 
      Caption         =   "Create Group"
      Height          =   345
      Left            =   120
      TabIndex        =   18
      Top             =   5880
      Width           =   1155
   End
   Begin VB.CommandButton cmdGetLinkInfo 
      Caption         =   "Get .lnk info... "
      Height          =   345
      Left            =   120
      TabIndex        =   17
      Top             =   6480
      Width           =   1155
   End
   Begin VB.ComboBox cmbSysFolders 
      Height          =   315
      Left            =   3120
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   16
      Top             =   90
      Width           =   5865
   End
   Begin VB.CommandButton cmdGetPath 
      Caption         =   "Get Sys Path"
      Height          =   345
      Left            =   -1200
      TabIndex        =   15
      Top             =   120
      Visible         =   0   'False
      Width           =   1155
   End
   Begin VB.TextBox txtShowCmd 
      Height          =   285
      Left            =   3120
      TabIndex        =   13
      Top             =   3975
      Width           =   585
   End
   Begin VB.TextBox txtCmdArgs 
      Height          =   285
      Left            =   3120
      TabIndex        =   11
      Top             =   2100
      Width           =   5865
   End
   Begin VB.TextBox txtIconIndex 
      Height          =   285
      Left            =   3120
      TabIndex        =   9
      Top             =   3480
      Width           =   585
   End
   Begin VB.TextBox txtIconFile 
      Height          =   285
      Left            =   3120
      TabIndex        =   7
      Top             =   3000
      Width           =   5865
   End
   Begin VB.TextBox txtWorkDir 
      Height          =   285
      Left            =   3120
      TabIndex        =   5
      Top             =   1725
      Width           =   5865
   End
   Begin VB.TextBox txtExeName 
      Height          =   285
      Left            =   3120
      TabIndex        =   3
      Top             =   1335
      Width           =   5865
   End
   Begin VB.TextBox txtLinkName 
      Height          =   285
      Left            =   3120
      TabIndex        =   1
      Top             =   960
      Width           =   5865
   End
   Begin VB.CommandButton cmdCreateLink 
      Caption         =   "Create Shortcut and Close"
      Height          =   585
      Left            =   120
      TabIndex        =   0
      Top             =   3600
      Width           =   1515
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Icon Index = 888"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   10
      Left            =   -240
      TabIndex        =   37
      Top             =   1080
      Width           =   1905
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      Caption         =   "Not used in BJ's Basic Calender. Code has not been deleted"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   120
      TabIndex        =   34
      Top             =   5280
      Width           =   9135
   End
   Begin VB.Label Label6 
      Caption         =   "Create shortcut in:"
      Height          =   255
      Left            =   1800
      TabIndex        =   33
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label Label5 
      Caption         =   "Current Path:"
      Height          =   255
      Left            =   2160
      TabIndex        =   31
      Top             =   600
      Width           =   975
   End
   Begin VB.Label Label4 
      Caption         =   "URL Target:"
      Height          =   255
      Left            =   2160
      TabIndex        =   30
      Top             =   7560
      Width           =   975
   End
   Begin VB.Label Label3 
      Caption         =   "URL Name:"
      Height          =   255
      Left            =   2160
      TabIndex        =   27
      Top             =   7200
      Width           =   855
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000004&
      X1              =   9120
      X2              =   240
      Y1              =   7080
      Y2              =   7080
   End
   Begin VB.Line Line1 
      BorderColor     =   &H8000000C&
      BorderWidth     =   2
      X1              =   9120
      X2              =   240
      Y1              =   7080
      Y2              =   7080
   End
   Begin VB.Label Label2 
      Caption         =   "3 = Miximum:     5 = Normal:     7 = Minimum"
      Height          =   255
      Left            =   3840
      TabIndex        =   25
      Top             =   3960
      Width           =   5175
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Comment:"
      Height          =   195
      Index           =   9
      Left            =   2280
      TabIndex        =   22
      Top             =   2520
      Width           =   705
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Shortcut Key:"
      Height          =   195
      Index           =   8
      Left            =   1920
      TabIndex        =   21
      Top             =   6240
      Width           =   960
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Start Menu Group:"
      Height          =   195
      Index           =   7
      Left            =   1650
      TabIndex        =   19
      Top             =   6720
      Width           =   1305
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Show Command:"
      Height          =   195
      Index           =   6
      Left            =   1860
      TabIndex        =   14
      Top             =   4020
      Width           =   1200
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Cmd Arguments:"
      Height          =   195
      Index           =   5
      Left            =   1890
      TabIndex        =   12
      Top             =   2160
      Width           =   1155
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Icon Index:"
      Height          =   195
      Index           =   4
      Left            =   2280
      TabIndex        =   10
      Top             =   3600
      Width           =   795
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Icon FileName:"
      Height          =   195
      Index           =   3
      Left            =   1965
      TabIndex        =   8
      Top             =   3120
      Width           =   1065
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Working Directory:"
      Height          =   195
      Index           =   2
      Left            =   1740
      TabIndex        =   6
      Top             =   1785
      Width           =   1320
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Exe Name:"
      Height          =   195
      Index           =   1
      Left            =   2280
      TabIndex        =   4
      Top             =   1395
      Width           =   780
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Link Name:"
      Height          =   195
      Index           =   0
      Left            =   2250
      TabIndex        =   2
      Top             =   1020
      Width           =   810
   End
End
Attribute VB_Name = "frmIcon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lngIcon, FB
Private Sub cmbSysFolders_Click()
Dim i As Integer
cmdGetPath_Click
cmbSysFolders.List(i) = cmbSysFolders.List(i)
End Sub

'---------------------------------------------------------------
Private Sub cmdCreateGroup_Click()
'---------------------------------------------------------------
    MkDir txtProgramGroup.Text                          ' Create Start Menu Program Group...
'---------------------------------------------------------------
End Sub
'---------------------------------------------------------------

'---------------------------------------------------------------
Private Sub cmdCreateLink_Click()
'---------------------------------------------------------------
    Dim sLnk As cShellLink                                ' ShellLink Variable
'---------------------------------------------------------------
    Set sLnk = New cShellLink                           ' Create ShellLink Instance
    
    sLnk.CreateShellLink txtPath & "\" & txtLinkName, txtExeName, txtWorkDir, txtCmdArgs, txtLnkDesc, txtIconFile, CLng(txtIconIndex), CLng(txtShowCmd)        ' Create a ShellLink (ShortCut)
    
    Set sLnk = Nothing                                  ' Destroy object reference
'---------------------------------------------------------------
End Sub
'---------------------------------------------------------------

'---------------------------------------------------------------
Private Sub cmdGetLinkInfo_Click()
'---------------------------------------------------------------
    Dim sLnk As cShellLink                              ' ShellLink class variable
    Dim LnkFile As String                               ' Link file name
    Dim ExeFile As String                               ' Link - Exe file name
    Dim WorkDir As String                               '      - Working directory
    Dim ExeArgs As String                               '      - Command line arguments
    Dim IconFile As String                              '      - Icon File name
    Dim IconIdx As Long                                 '      - Icon Index
    Dim ShowCmd As Long                                 '      - Program start state...
'---------------------------------------------------------------
    Set sLnk = New cShellLink                           ' Create new Explorer IShellLink Instance
    
    LnkFile = txtLinkName.Text                          ' Get link file name
    txtExeName.Text = ""                                ' Clear output variables...
    txtWorkDir.Text = ""
    txtCmdArgs.Text = ""
    txtIconFile.Text = ""
    txtIconIndex.Text = ""
    txtShowCmd.Text = ""
    
    sLnk.GetShellLinkInfo LnkFile, ExeFile, WorkDir, ExeArgs, IconFile, IconIdx, ShowCmd                   ' Get Info for shortcut file...
                        
    txtLinkName.Text = LnkFile                          ' Display output...
    txtExeName.Text = ExeFile
    txtWorkDir.Text = WorkDir
    txtCmdArgs.Text = ExeArgs
    txtIconFile.Text = IconFile
    txtIconIndex.Text = Val(IconIdx)
    txtShowCmd.Text = Val(ShowCmd)
    
    Set sLnk = Nothing                                  ' Destroy object reference...
'---------------------------------------------------------------
End Sub
'---------------------------------------------------------------

'---------------------------------------------------------------
Private Sub cmdGetPath_Click()
'---------------------------------------------------------------
    Dim rc As Long                                      ' return code
    Dim sLnk As cShellLink                              ' ShellLink class object
    Dim sfPath As String                                ' System folder path
    Dim Id As Long                                      ' ID of System folder...
'---------------------------------------------------------------
    ' Create instance of Explorer's IShellLink Interface Base Class
    Set sLnk = New cShellLink
    
    Id = cmbSysFolders.ItemData(cmbSysFolders.ListIndex)  ' Get ID from combo box
    If sLnk.GetSystemFolderPath(Me.hwnd, Id, sfPath) Then ' Get system folder path from id
        SetDefaults sfPath                                ' Update UI with new path
    End If
    
    Set sLnk = Nothing                                  ' Destroy object reference
'---------------------------------------------------------------
End Sub
'---------------------------------------------------------------

Private Sub cmdURL_Click()
'---------------------------------------------------------------
    Dim sUrl As cShellLink                              ' ShellLink class variable
    Dim UrlFile As String                               ' Link file name
'---------------------------------------------------------------
    Set sUrl = New cShellLink                           ' Create ShellLink Instance
    
    sUrl.CreateInternetLink txtURL.Text, txtURLTarget.Text        ' Create a ShellURLLink (URL ShortCut)
    
    Set sUrl = Nothing                                  ' Destroy object reference
'---------------------------------------------------------------

End Sub

'---------------------------------------------------------------
Private Sub Form_Load()
    Dim AppPath As String                                           ' Current Application path
    AppPath = App.Path                                              ' Get current path
    
    If (Right$(AppPath, 1) <> "\") Then AppPath = AppPath & "\" ' Fix application path if necessary
'---------------------------------------------------------------
    SetDefaults (App.Path & "\")                    ' Update UI with current application path
    
    With cmbSysFolders                              ' Add ID's for system folders to combo box...
        .AddItem "DESKTOP"
        .ItemData(.NewIndex) = 0
        .AddItem "Programs"
        .ItemData(.NewIndex) = &H2
        .AddItem "Control"
        .ItemData(.NewIndex) = &H3
        .AddItem "Printers"
        .ItemData(.NewIndex) = &H4
        .AddItem "Personal"
        .ItemData(.NewIndex) = &H5
        .AddItem "Favorites"
        .ItemData(.NewIndex) = &H6
        .AddItem "Startup"
        .ItemData(.NewIndex) = &H7
        .AddItem "Recent"
        .ItemData(.NewIndex) = &H8
        .AddItem "Send to..."
        .ItemData(.NewIndex) = &H9
        .AddItem "Recycle Bin"
        .ItemData(.NewIndex) = &HA
        .AddItem "Start Menu"
        .ItemData(.NewIndex) = &HB
        .AddItem "Desktop Directory"
        .ItemData(.NewIndex) = &H10
        .AddItem "DRIVES"
        .ItemData(.NewIndex) = &H11
        .AddItem "Network"
        .ItemData(.NewIndex) = &H12
        .AddItem "Net Hood"
        .ItemData(.NewIndex) = &H13
        .AddItem "Fonts"
        .ItemData(.NewIndex) = &H14
        .AddItem "Templates"
        .ItemData(.NewIndex) = &H15
        .AddItem "Common Start Menu"
        .ItemData(.NewIndex) = &H16
        .AddItem "Common Programs"
        .ItemData(.NewIndex) = &H17
        .AddItem "Common Startup"
        .ItemData(.NewIndex) = &H18
        .AddItem "Common Desktop Directory"
        .ItemData(.NewIndex) = &H19
        .AddItem "App Data"
        .ItemData(.NewIndex) = &H1A
        .AddItem "Print Hood"
        .ItemData(.NewIndex) = &H1B
        
        
        
        .ListIndex = 6
    End With
    
                lngIcon = ExtractIcon(App.hInstance, AppPath & App.EXEName & ".exe", 0)
            If lngIcon = 0 Then
                MsgBox "Error loading icons into menu from (" & AppPath & App.EXEName & ".exe" & ")"
                PicIconIndex.Max = 0
                txtExeName.Text = ""
                txtPath.Text = ""
                pctIcon.Cls
            Else
                pctIcon.Cls
                PicIconIndex.Value = 0
                lngIcon = ExtractIcon(App.hInstance, AppPath & App.EXEName & ".exe", 0)
                DrawIcon pctIcon.hDC, 0, 0, lngIcon
                
                FB = 0
                Do
                    lngIcon = ExtractIcon(App.hInstance, AppPath & App.EXEName & ".exe", FB)
                    FB = FB + 1
                    If lngIcon = 0 Then
                        PicIconIndex.Max = FB - 2
                        GoTo FinaliseError
                    End If
                Loop
FinaliseError:
            End If

'---------------------------------------------------------------
End Sub
'---------------------------------------------------------------

'---------------------------------------------------------------
Private Sub SetDefaults(pth As String)
'---------------------------------------------------------------
    Dim AppPath As String                                           ' Current Application path
    Dim rc As Long                                      ' return code
    Dim sLnk As cShellLink                              ' ShellLink class object
    Dim sfPath As String                                ' System folder path
    Dim Id As Long                                      ' ID of System folder...
'---------------------------------------------------------------
    ' Create instance of Explorer's IShellLink Interface Base Class
    Set sLnk = New cShellLink
'---------------------------------------------------------------
    AppPath = App.Path                                              ' Get current path
    
    If (Right$(AppPath, 1) <> "\") Then AppPath = AppPath & "\" ' Fix application path if necessary
    If (Right$(pth, 1) <> "\") Then pth = pth & "\"         ' Fix path if necessary
     
  
    
    If sLnk.GetSystemFolderPath(Me.hwnd, &HB, sfPath) Then ' Get system folder path from id
  
    txtPath.Text = pth                                       ' Create a full path name for link file
    txtLinkName.Text = ".lnk"                             ' Create a full path name for link file
    txtExeName.Text = ".exe"                                    ' Create a full path name for applicaton exe name
    txtWorkDir.Text = ""                                            ' Set default working directory
    txtLnkDesc.Text = ""                                            ' Set comment
    txtwHotKey.Text = "CTRL + ALT + "                         ' Set hot key
    txtIconFile.Text = txtExeName.Text                          ' Set default IconFile name to default exename
    txtIconIndex.Text = CStr(0)                                     ' Set default Icon Index val
    txtShowCmd.Text = CStr(5)                                   ' set default showcommand val
    txtProgramGroup.Text = sfPath & "\"                             ' Set default Program group name
    txtURL.Text = txtPath & ".url"                                      ' Create a full path name for URL file
    txtURLTarget.Text = "www."                                          ' Create a full path name for the web site
'---------------------------------------------------------------
    
    End If
      

End Sub
'---------------------------------------------------------------
Private Sub txtExeName_Change()
txtIconFile.Text = txtExeName
End Sub


Private Sub PicIconIndex_Change()
    Dim AppPath As String                                           ' Current Application path
    AppPath = App.Path                                              ' Get current path
    
    If (Right$(AppPath, 1) <> "\") Then AppPath = AppPath & "\" ' Fix application path if necessary
     
    lngIcon = ExtractIcon(App.hInstance, AppPath & App.EXEName & ".exe", PicIconIndex.Value)
    If lngIcon = 0 Then
        MsgBox "Error loading icons into menu from (" & AppPath & App.EXEName & ".exe" & ")"
    Else
        pctIcon.Cls
        lngIcon = ExtractIcon(App.hInstance, AppPath & App.EXEName & ".exe", PicIconIndex.Value)
        DrawIcon pctIcon.hDC, 0, 0, lngIcon
    End If
End Sub

