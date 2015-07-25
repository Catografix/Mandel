VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.MDIForm frmMain 
   BackColor       =   &H8000000C&
   Caption         =   "MandelProject1"
   ClientHeight    =   3195
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   4680
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog dlgCommonDialog 
      Left            =   2160
      Top             =   1320
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin ComctlLib.StatusBar sbStatusBar 
      Align           =   2  'Align Bottom
      Height          =   270
      Left            =   0
      TabIndex        =   0
      Top             =   2925
      Width           =   4680
      _ExtentX        =   8255
      _ExtentY        =   476
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   3
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   1
            Object.Width           =   2699
            Text            =   "Status"
            TextSave        =   "Status"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   6
            AutoSize        =   2
            TextSave        =   "11/8/98"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   5
            AutoSize        =   2
            TextSave        =   "5:36 PM"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileNew 
         Caption         =   "&New"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuFileOpen 
         Caption         =   "&Open"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuFileClose 
         Caption         =   "&Close"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuFileBar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileSave 
         Caption         =   "&Save"
         Enabled         =   0   'False
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuFileSaveAs 
         Caption         =   "Save &As..."
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuFileSaveAll 
         Caption         =   "Save A&ll"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileBar2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileProperties 
         Caption         =   "Propert&ies"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuFileBar3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFilePageSetup 
         Caption         =   "Page Set&up..."
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuFilePrintPreview 
         Caption         =   "Print Pre&view"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuFilePrint 
         Caption         =   "&Print..."
         Enabled         =   0   'False
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuFileBar4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileSend 
         Caption         =   "Sen&d..."
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuFileBar5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileMRU 
         Caption         =   ""
         Index           =   0
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileMRU 
         Caption         =   ""
         Index           =   1
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileMRU 
         Caption         =   ""
         Index           =   2
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileMRU 
         Caption         =   ""
         Index           =   3
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileBar6 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Enabled         =   0   'False
      Begin VB.Menu mnuEditUndo 
         Caption         =   "&Undo"
         Shortcut        =   ^Z
      End
      Begin VB.Menu mnuEditBar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditCut 
         Caption         =   "Cu&t"
         Shortcut        =   ^X
      End
      Begin VB.Menu mnuEditCopy 
         Caption         =   "&Copy"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuEditPaste 
         Caption         =   "&Paste"
         Shortcut        =   ^V
      End
      Begin VB.Menu mnuEditPasteSpecial 
         Caption         =   "Paste &Special..."
      End
   End
   Begin VB.Menu mnuWindow 
      Caption         =   "&Window"
      Enabled         =   0   'False
      WindowList      =   -1  'True
      Begin VB.Menu mnuWindowNewWindow 
         Caption         =   "&New Window"
      End
      Begin VB.Menu mnuWindowBar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuWindowCascade 
         Caption         =   "&Cascade"
      End
      Begin VB.Menu mnuWindowTileHorizontal 
         Caption         =   "Tile &Horizontal"
      End
      Begin VB.Menu mnuWindowTileVertical 
         Caption         =   "Tile &Vertical"
      End
      Begin VB.Menu mnuWindowArrangeIcons 
         Caption         =   "&Arrange Icons"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpContents 
         Caption         =   "&Contents"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuHelpSearch 
         Caption         =   "&Search For Help On..."
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuHelpBar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "&About MandelProject1..."
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function OSWinHelp% Lib "user32" Alias "WinHelpA" (ByVal hwnd&, ByVal HelpFile$, ByVal wCommand%, dwData As Any)
Public sFile As String
Private Sub MDIForm_Load()
    Me.Left = GetSetting(App.Title, "Settings", "MainLeft", 1000)
    Me.Top = GetSetting(App.Title, "Settings", "MainTop", 1000)
    Me.Width = GetSetting(App.Title, "Settings", "MainWidth", 6500)
    Me.Height = GetSetting(App.Title, "Settings", "MainHeight", 6500)
    Close
End Sub


Private Sub LoadNewDoc()
    Static lDocumentCount As Long
    Dim frmD As frmDocument


    lDocumentCount = lDocumentCount + 1
    Set frmD = New frmDocument
    frmD.Caption = "Document " & lDocumentCount
    frmD.Show
End Sub


Private Sub MDIForm_Unload(Cancel As Integer)
    If Me.WindowState <> vbMinimized Then
        SaveSetting App.Title, "Settings", "MainLeft", Me.Left
        SaveSetting App.Title, "Settings", "MainTop", Me.Top
        SaveSetting App.Title, "Settings", "MainWidth", Me.Width
        SaveSetting App.Title, "Settings", "MainHeight", Me.Height
    End If
    Close
    Unload Me
End Sub



Private Sub mnuHelpAbout_Click()
    frmAbout.Show vbModal, Me
End Sub



Private Sub tbToolBar_ButtonClick(ByVal button As ComctlLib.button)


    Select Case button.Key


        Case "New"
            LoadNewDoc
        Case "New"
            mnuFileNew_Click
        Case "Open"
            mnuFileOpen_Click
        Case "Save"
            mnuFileSave_Click
        Case "Print"
            mnuFilePrint_Click
        Case "Cut"
            mnuEditCut_Click
        Case "Copy"
            mnuEditCopy_Click
        Case "Paste"
            mnuEditPaste_Click
        Case "Bold"
            'To Do
            MsgBox "Bold Code goes here!"
        Case "Italic"
            'To Do
            MsgBox "Italic Code goes here!"
        Case "Underline"
            'To Do
            MsgBox "Underline Code goes here!"
        Case "Left"
            'To Do
            MsgBox "Left Code goes here!"
        Case "Center"
            'To Do
            MsgBox "Center Code goes here!"
        Case "Right"
            'To Do
            MsgBox "Right Code goes here!"
    End Select
End Sub







Private Sub mnuHelpContents_Click()
    

    Dim nRet As Integer


    'if there is no helpfile for this project display a message to the user
    'you can set the HelpFile for your application in the
    'Project Properties dialog
    If Len(App.HelpFile) = 0 Then
        MsgBox "Unable to display Help Contents. There is no Help associated with this project.", vbInformation, Me.Caption
    Else
        On Error Resume Next
        nRet = OSWinHelp(Me.hwnd, App.HelpFile, 3, 0)
        If Err Then
            MsgBox Err.Description
        End If
    End If
End Sub


Private Sub mnuHelpSearch_Click()
    

    Dim nRet As Integer


    'if there is no helpfile for this project display a message to the user
    'you can set the HelpFile for your application in the
    'Project Properties dialog
    If Len(App.HelpFile) = 0 Then
        MsgBox "Unable to display Help Contents. There is no Help associated with this project.", vbInformation, Me.Caption
    Else
        On Error Resume Next
        nRet = OSWinHelp(Me.hwnd, App.HelpFile, 261, 0)
        If Err Then
            MsgBox Err.Description
        End If
    End If
End Sub


Private Sub mnuWindowArrangeIcons_Click()
    Me.Arrange vbArrangeIcons
End Sub


Private Sub mnuWindowCascade_Click()
    Me.Arrange vbCascade
End Sub


Private Sub mnuWindowNewWindow_Click()
    'To Do
    MsgBox "New Window Code goes here!"
End Sub


Private Sub mnuWindowTileHorizontal_Click()
    Me.Arrange vbTileHorizontal
End Sub


Private Sub mnuWindowTileVertical_Click()
    Me.Arrange vbTileVertical
End Sub

Private Sub mnuEditCopy_Click()
    'To Do
    MsgBox "Copy Code goes here!"
End Sub


Private Sub mnuEditCut_Click()
    'To Do
    MsgBox "Cut Code goes here!"
End Sub


Private Sub mnuEditPaste_Click()
    'To Do
    MsgBox "Paste Code goes here!"
End Sub


Private Sub mnuEditPasteSpecial_Click()
    'To Do
    MsgBox "Paste Special Code goes here!"
End Sub


Private Sub mnuEditUndo_Click()
    'To Do
    MsgBox "Undo Code goes here!"
End Sub

Private Sub mnuFileOpen_Click()

On Error GoTo DialogError

    With dlgCommonDialog
        .CancelError = True
        'set the flags and attributes of the
        'common dialog control
        .Filter = "All Files (*.*)|*.*"
        .ShowOpen
        If Len(.filename) = 0 Then
            Exit Sub
        End If
        sFile = .filename
    End With
    Close
    Dim Picture1 As New Form6
    Picture1.sFilename = sFile
    Picture1.Caption = sFile
    Load Picture1
    Picture1.Show
    Exit Sub
DialogError:
    MsgBox Str(Err.Number) & ":" & Err.Description, , "Error"
End Sub


Private Sub mnuFileSave_Click()
    If MsgBox("Replace Color Data?", vbOKCancel) = vbCancel Then Exit Sub
 With dlgCommonDialog
        .CancelError = True
        'set the flags and attributes of the
        'common dialog control
        .Filter = "All Files (*.*)|*.*"
        .ShowSave
        If Len(.filename) = 0 Then
            Exit Sub
        End If
        sFile = .filename
    End With
    'Form6.replace (sFile)
    Exit Sub
DialogError:
    MsgBox Str(Err.Number) & ":" & Err.Description, , "Error"
End Sub


Private Sub mnuFileSaveAs_Click()
    'To Do
    'Setup the common dialog control
    'prior to calling ShowSave
    dlgCommonDialog.ShowSave
End Sub


Private Sub mnuFileSaveAll_Click()
    'To Do
    MsgBox "Save All Code goes here!"
End Sub


Private Sub mnuFileProperties_Click()
    'To Do
    MsgBox "Properties Code goes here!"
End Sub


Private Sub mnuFilePageSetup_Click()
    dlgCommonDialog.ShowPrinter
End Sub


Private Sub mnuFilePrintPreview_Click()
    'To Do
    MsgBox "Print Preview Code goes here!"
End Sub


Private Sub mnuFilePrint_Click()
    'To Do
    MsgBox "Print Code goes here!"
End Sub


Private Sub mnuFileSend_Click()
    'To Do
    MsgBox "Send Code goes here!"
End Sub


Private Sub mnuFileMRU_Click(Index As Integer)
    'To Do
    MsgBox "MRU Code goes here!"
End Sub


Private Sub mnuFileExit_Click()
    'unload the form
    Close
    Unload Me
End Sub

Private Sub mnuFileNew_Click()
    Form1.Caption = "Data Input"
    Form1.Show
End Sub



