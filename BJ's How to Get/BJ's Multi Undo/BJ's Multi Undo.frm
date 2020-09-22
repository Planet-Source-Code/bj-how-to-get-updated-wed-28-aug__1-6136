VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmMultiple 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "BJ's How to Get... Multi Undo."
   ClientHeight    =   3855
   ClientLeft      =   2400
   ClientTop       =   1890
   ClientWidth     =   6675
   ClipControls    =   0   'False
   Icon            =   "BJ's Multi Undo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3855
   ScaleWidth      =   6675
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdEMail 
      Caption         =   "E-Mail"
      Height          =   495
      Left            =   4080
      TabIndex        =   11
      Top             =   3240
      Width           =   1215
   End
   Begin VB.CommandButton cmdAbout 
      Caption         =   "About"
      Height          =   495
      Left            =   2760
      TabIndex        =   10
      Top             =   3240
      Width           =   1215
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   495
      Left            =   5400
      TabIndex        =   9
      Top             =   3240
      Width           =   1215
   End
   Begin VB.CommandButton cmdSelectAll 
      Caption         =   "&Select All"
      Height          =   495
      Left            =   1440
      TabIndex        =   8
      Top             =   3240
      Width           =   1215
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "&Delete"
      Height          =   495
      Left            =   120
      TabIndex        =   7
      Top             =   3240
      Width           =   1215
   End
   Begin VB.CommandButton cmdPaste 
      Caption         =   "&Paste"
      Height          =   495
      Left            =   5400
      TabIndex        =   6
      Top             =   2640
      Width           =   1215
   End
   Begin VB.CommandButton cmdCopy 
      Caption         =   "C&opy"
      Height          =   495
      Left            =   4080
      TabIndex        =   5
      Top             =   2640
      Width           =   1215
   End
   Begin VB.CommandButton cmdCut 
      Caption         =   "&Cut"
      Height          =   495
      Left            =   2760
      TabIndex        =   4
      Top             =   2640
      Width           =   1215
   End
   Begin VB.CommandButton cmdDummy 
      Height          =   315
      Left            =   7980
      TabIndex        =   2
      Top             =   4860
      Width           =   1215
   End
   Begin VB.CommandButton cmdRedo 
      Caption         =   "&Redo"
      Height          =   495
      Left            =   1440
      TabIndex        =   1
      Top             =   2640
      Width           =   1215
   End
   Begin VB.CommandButton cmdUndo 
      Caption         =   "&Undo"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   2640
      Width           =   1215
   End
   Begin RichTextLib.RichTextBox txtEdit 
      Height          =   2535
      Left            =   120
      TabIndex        =   3
      Top             =   0
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   4471
      _Version        =   393217
      TextRTF         =   $"BJ's Multi Undo.frx":014A
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Visible         =   0   'False
      Begin VB.Menu mnuUndo 
         Caption         =   "&Undo"
         Shortcut        =   ^Z
      End
      Begin VB.Menu mnuRedo 
         Caption         =   "&Redo"
         Shortcut        =   ^Y
      End
      Begin VB.Menu mnuSep0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCut 
         Caption         =   "&Cut"
         Shortcut        =   ^X
      End
      Begin VB.Menu mnuCopy 
         Caption         =   "&Copy"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuPaste 
         Caption         =   "&Paste"
         Shortcut        =   ^V
      End
      Begin VB.Menu mnuDelete 
         Caption         =   "&Delete"
      End
      Begin VB.Menu mnuSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSelectAll 
         Caption         =   "Select &All"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
         Shortcut        =   {F4}
      End
   End
End
Attribute VB_Name = "frmMultiple"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
    (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, _
    ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private trapUndo As Boolean           'flag to indicate whether actions should be trapped
Private UndoStack As New Collection   'collection of undo elements
Private RedoStack As New Collection   'collection of redo elements
Private Declare Function ShellAbout Lib "shell32.dll" Alias "ShellAboutA" _
(ByVal hwnd As Long, ByVal szApp As String, ByVal szOtherStuff As String, _
ByVal hIcon As Long) As Long



Public Sub AboutBox(hwnd As Long)
'-----------------------------------------------------------------
    ' Show help about dialog...
    ShellAbout hwnd, App.EXEName, _
               vbCrLf & "BJ. E-Mail me: bryce3@bigpond.com", frmMultiple.Icon
'-----------------------------------------------------------------
End Sub

Private Sub cmdAbout_Click()
AboutBox Me.hwnd
End Sub

Private Sub cmdCopy_Click()
mnuCopy_Click
End Sub

Private Sub cmdCut_Click()
mnuCut_Click
End Sub

Private Sub cmdDelete_Click()
mnuDelete_Click
End Sub

Private Sub cmdEMail_Click()
    ShellExecute 0, "Open", "mailto:bryce3@bigpond.com?subject=Info about BJ's How to Get... Multi Undo.&body=Type here the message that you want to send me.", "", "", vbMaximizedFocus

End Sub

Private Sub cmdExit_Click()
mnuExit_Click
End Sub

Private Sub cmdPaste_Click()
mnuPaste_Click
End Sub

Private Sub cmdRedo_Click()
    Redo
End Sub

Private Sub cmdSelectAll_Click()
mnuSelectAll_Click
End Sub

Private Sub cmdUndo_Click()
    Undo
End Sub

Private Sub Form_Load()
If App.PrevInstance = True Then
End
End If
CreateKey "HKEY_CURRENT_USER\Software\BJ\BJ's How to Get..."
SetStringValue "HKEY_CURRENT_USER\Software\BJ\BJ's How to Get...", "Multi Undo", App.EXEName & "- Last time run was " & " - " & Now
    trapUndo = True     'Enable Undo Trapping
    txtEdit_Change      'Initialize First Undo
    txtEdit_SelChange   'Initialize Menus
    Show
    DoEvents
End Sub

Private Sub mnuCopy_Click()
    Clipboard.SetText txtEdit.SelText, 1
End Sub

Private Sub mnuCut_Click()
    Clipboard.SetText txtEdit.SelText, 1
    txtEdit.SelText = ""
End Sub

Private Sub mnuDelete_Click()
    txtEdit.SelText = ""
End Sub

Private Sub mnuExit_Click()
If App.EXEName = "BJ's Multi Undo" Then
End
Else
Unload frmMultiple
frmHowtoGet.Show
End If
End Sub

Private Sub mnuPaste_Click()
    txtEdit.SelText = ""                    'This step is crucial!!! for undoing actions
    txtEdit.SelText = Clipboard.GetText(1)
End Sub

Private Sub mnuRedo_Click()
    cmdRedo_Click
End Sub

Private Sub mnuSelectAll_Click()
    txtEdit.SelStart = 0
    txtEdit.SelLength = Len(txtEdit.Text)
End Sub

Private Sub mnuUndo_Click()
    cmdUndo_Click
End Sub

Private Sub txtEdit_Change()
    If Not trapUndo Then Exit Sub 'because trapping is disabled

    Dim newElement As New clsMultiUndo   'create new undo element
    Dim c%, l&

    'remove all redo items because of the change
    For c% = 1 To RedoStack.Count
        RedoStack.Remove 1
    Next c%

    'set the values of the new element
    newElement.SelStart = txtEdit.SelStart
    newElement.TextLen = Len(txtEdit.Text)
    newElement.Text = txtEdit.Text

    'add it to the undo stack
    UndoStack.Add Item:=newElement
    'enable controls accordingly
    EnableControls
End Sub

Private Sub txtEdit_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift = 2 Then 'a control event (Ctrl + C, Ctrl + Z), etc.
            KeyCode = 0
    End If
End Sub

Private Sub txtEdit_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton Then 'do the popup menu
        PopupMenu mnuEdit
    End If
End Sub

Private Sub txtEdit_SelChange()
Dim ln&
    If Not trapUndo Then Exit Sub
    ln& = txtEdit.SelLength
    mnuCut.Enabled = ln&    'disabled if length of selected text is 0
    mnuCopy.Enabled = ln&   'disabled if length of selected text is 0
    mnuPaste.Enabled = Len(Clipboard.GetText(1)) 'disabled if length of clipboard text is 0
    mnuDelete.Enabled = ln&  'disabled if length of selected text is 0
    mnuSelectAll.Enabled = CBool(Len(txtEdit.Text)) 'disabled if length of textbox's text is 0
End Sub

Private Sub EnableControls()
    cmdUndo.Enabled = UndoStack.Count > 1
    cmdRedo.Enabled = RedoStack.Count > 0
    mnuUndo.Enabled = cmdUndo.Enabled
    mnuRedo.Enabled = cmdRedo.Enabled
    txtEdit_SelChange
End Sub

Public Function Change(ByVal lParam1 As String, ByVal lParam2 As String, startSearch As Long) As String
Dim tempParam$
Dim d&
    If Len(lParam1) > Len(lParam2) Then 'swap
        tempParam$ = lParam1
        lParam1 = lParam2
        lParam2 = tempParam$
    End If
    d& = Len(lParam2) - Len(lParam1)
    Change = Mid(lParam2, startSearch - d&, d&)
End Function

Public Sub Undo()
Dim chg$, X&
Dim DeleteFlag As Boolean 'flag as to whether or not to delete text or append text
Dim objElement As Object, objElement2 As Object
    If UndoStack.Count > 1 And trapUndo Then 'we can proceed
        trapUndo = False
        DeleteFlag = UndoStack(UndoStack.Count - 1).TextLen < UndoStack(UndoStack.Count).TextLen
        If DeleteFlag Then  'delete some text
            cmdDummy.SetFocus   'change focus of form
            X& = SendMessage(txtEdit.hwnd, EM_HIDESELECTION, 1&, 1&)
            Set objElement = UndoStack(UndoStack.Count)
            Set objElement2 = UndoStack(UndoStack.Count - 1)
            txtEdit.SelStart = objElement.SelStart - (objElement.TextLen - objElement2.TextLen)
            txtEdit.SelLength = objElement.TextLen - objElement2.TextLen
            txtEdit.SelText = ""
            X& = SendMessage(txtEdit.hwnd, EM_HIDESELECTION, 0&, 0&)
        Else 'append something
            Set objElement = UndoStack(UndoStack.Count - 1)
            Set objElement2 = UndoStack(UndoStack.Count)
            chg$ = Change(objElement.Text, objElement2.Text, _
                objElement2.SelStart + 1 + Abs(Len(objElement.Text) - Len(objElement2.Text)))
            txtEdit.SelStart = objElement2.SelStart
            txtEdit.SelLength = 0
            txtEdit.SelText = chg$
            txtEdit.SelStart = objElement2.SelStart
            If Len(chg$) > 1 And chg$ <> vbCrLf Then
                txtEdit.SelLength = Len(chg$)
            Else
                txtEdit.SelStart = txtEdit.SelStart + Len(chg$)
            End If
        End If
        RedoStack.Add Item:=UndoStack(UndoStack.Count)
        UndoStack.Remove UndoStack.Count
    End If
    EnableControls
    trapUndo = True
    txtEdit.SetFocus
End Sub

Public Sub Redo()
Dim chg$
Dim DeleteFlag As Boolean 'flag as to whether or not to delete text or append text
Dim objElement As Object
    If RedoStack.Count > 0 And trapUndo Then
        trapUndo = False
        DeleteFlag = RedoStack(RedoStack.Count).TextLen < Len(txtEdit.Text)
        If DeleteFlag Then  'delete last item
            Set objElement = RedoStack(RedoStack.Count)
            txtEdit.SelStart = objElement.SelStart
            txtEdit.SelLength = Len(txtEdit.Text) - objElement.TextLen
            txtEdit.SelText = ""
        Else 'append something
            Set objElement = RedoStack(RedoStack.Count)
            chg$ = Change(txtEdit.Text, objElement.Text, objElement.SelStart + 1)
            txtEdit.SelStart = objElement.SelStart - Len(chg$)
            txtEdit.SelLength = 0
            txtEdit.SelText = chg$
            txtEdit.SelStart = objElement.SelStart - Len(chg$)
            If Len(chg$) > 1 And chg$ <> vbCrLf Then
                txtEdit.SelLength = Len(chg$)
            Else
                txtEdit.SelStart = txtEdit.SelStart + Len(chg$)
            End If
        End If
        UndoStack.Add Item:=objElement
        RedoStack.Remove RedoStack.Count
    End If
    EnableControls
    trapUndo = True
    txtEdit.SetFocus
End Sub
