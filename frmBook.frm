VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmBook 
   Caption         =   "Book"
   ClientHeight    =   3195
   ClientLeft      =   225
   ClientTop       =   825
   ClientWidth     =   4680
   Icon            =   "frmBook.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog cdlSave 
      Left            =   2160
      Top             =   1320
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComDlg.CommonDialog cdlgFont 
      Left            =   2160
      Top             =   1320
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Default         =   -1  'True
      Height          =   495
      Left            =   1800
      TabIndex        =   1
      Top             =   2520
      Width           =   1215
   End
   Begin RichTextLib.RichTextBox rtfMain 
      Height          =   2295
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   4048
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"frmBook.frx":0442
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   13.5
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileOpen 
         Caption         =   "&Import a book"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuFileSaveas 
         Caption         =   "&Export a book"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuEditCut 
         Caption         =   "C&ut"
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
      Begin VB.Menu mnuEditFind 
         Caption         =   "&Find"
         Shortcut        =   ^F
      End
      Begin VB.Menu mnuEditFindnext 
         Caption         =   "Find &next"
         Shortcut        =   {F3}
      End
      Begin VB.Menu mnuEditSelectall 
         Caption         =   "Select &all"
         Shortcut        =   ^A
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "&View"
      Begin VB.Menu mnuViewWhat 
         Caption         =   "&Catalogue"
         Index           =   0
         Shortcut        =   ^L
      End
      Begin VB.Menu mnuViewWhat 
         Caption         =   "&Book"
         Index           =   1
         Shortcut        =   ^B
      End
      Begin VB.Menu mnuViewWhat 
         Caption         =   "&Downloaded books"
         Index           =   2
         Shortcut        =   ^D
      End
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "&Options"
      Begin VB.Menu mnuOptionsFont 
         Caption         =   "&Font"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpManual 
         Caption         =   "&Manual"
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuHelpLegal 
         Caption         =   "&Legal"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "&About"
      End
   End
End
Attribute VB_Name = "frmBook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mFind As String

Private Sub cmdOK_Click()
    On Error Resume Next
    Call Me.Hide
End Sub

Private Sub Form_Initialize()
    On Error Resume Next
    Call modXPStyle.InitCommonControlsVB
End Sub

Private Sub Form_Load()
    On Error Resume Next
    Call modI18N.ApplyUILanguageToThisForm(Me)
    Call modLargeFonts.ApplySystemSettingsToForm(Me, , True)
    rtfMain.Font.bold = CBool(modPath.GetSettingIni(App.EXEName, "Font", "Bold", "True"))
    rtfMain.Font.name = modPath.GetSettingIni(App.EXEName, "Font", "Name", "Tahoma")
    rtfMain.Font.Charset = CInt(modPath.GetSettingIni(App.EXEName, "Font", "Charset", "0"))
    rtfMain.Font.Italic = CBool(modPath.GetSettingIni(App.EXEName, "Font", "Italic", "False"))
    rtfMain.Font.size = CInt(modPath.GetSettingIni(App.EXEName, "Font", "Size", "14"))
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    Me.Left = frmGutenberg.Left
    Me.Top = frmGutenberg.Top
    Me.Width = frmGutenberg.Width
    Me.Height = frmGutenberg.Height
    rtfMain.Left = 0
    rtfMain.Top = 0
    rtfMain.Height = Me.ScaleHeight - cmdOK.Height - GAP - GAP
    rtfMain.Width = Me.ScaleWidth
    cmdOK.Left = Me.ScaleWidth / 2 - cmdOK.Width / 2
    cmdOK.Top = Me.ScaleHeight - GAP - cmdOK.Height
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    'save current cursor position
    Call SaveCurrentCursorPosition
    If rtfMain.filename <> "" Then Call rtfMain.SaveFile(rtfMain.filename)
    Call modPath.SaveSettingIni(App.EXEName, "Font", "Bold", CStr(rtfMain.Font.bold))
    Call modPath.SaveSettingIni(App.EXEName, "Font", "Name", rtfMain.Font.name)
    Call modPath.SaveSettingIni(App.EXEName, "Font", "Charset", rtfMain.Font.Charset)
    Call modPath.GetSettingIni(App.EXEName, "Font", "Italic", rtfMain.Font.Italic)
    Call modPath.GetSettingIni(App.EXEName, "Font", "Size", rtfMain.Font.size)
End Sub

Private Sub mnuEditCopy_Click()
    On Error Resume Next
    Call Clipboard.Clear
    Call Clipboard.SetText(rtfMain.SelText)
End Sub

Private Sub mnuEditCut_Click()
    On Error Resume Next
    Call Clipboard.Clear
    Call Clipboard.SetText(rtfMain.SelText)
    rtfMain.SelText = ""
End Sub

Private Sub mnuEditFind_Click()
    On Error Resume Next
    Dim s As String
    s = InputBox("Find what:", "Find", mFind)
    If Len(s) > 0 Then
        mFind = s
        Call Find(s)
    End If
End Sub

Private Sub Find(what As String)
    On Error Resume Next
    Dim found As Long
    
    found = InStr(rtfMain.SelStart + 2, rtfMain.text, what, vbTextCompare)
    If found > 0 Then
        rtfMain.SelStart = found - 1
        rtfMain.SelLength = Len(what)
        rtfMain.SetFocus
    Else
        found = InStr(1, rtfMain.text, what, vbTextCompare)
        If found > 0 Then
            rtfMain.SelStart = found
            rtfMain.SelLength = Len(what)
            rtfMain.SetFocus
        Else
            MsgBox "Cannot find " & what, vbInformation
        End If
    End If
End Sub

Private Sub mnuEditFindnext_Click()
    On Error Resume Next
    Call Find(mFind)
End Sub

Private Sub mnuEditPaste_Click()
    On Error Resume Next
    Call SendKeys("^V")
End Sub

Private Sub mnuEditSelectall_Click()
    On Error Resume Next
    Call SendKeys("^A")
End Sub

Private Sub mnuFileExit_Click()
    On Error Resume Next
    Call Unload(Me)
End Sub

Private Sub mnuFileOpen_Click()
    On Error Resume Next
    Call Me.Hide
    Call frmGutenberg.mnuFileOpen_Click
End Sub

Private Sub mnuFileSaveas_Click()
    On Error Resume Next
    cdlSave.cancelError = True
    cdlSave.DefaultExt = "txt"
    cdlSave.DialogTitle = GetText("Save book as...")
    cdlSave.Filter = "Text files (*.txt)|*.txt;"
    cdlSave.Flags = cdlOFNOverwritePrompt + cdlOFNPathMustExist
    On Error GoTo cancelError:
    cdlSave.ShowSave
    On Error Resume Next
    If cdlSave.filename <> "" Then
        Call rtfMain.SaveFile(cdlSave.filename, rtfText)
    End If
    Exit Sub
cancelError:
    On Error Resume Next
End Sub

Private Sub mnuHelpAbout_Click()
    On Error Resume Next
    Call Me.Hide
    Call frmGutenberg.mnuHelpAbout_Click
End Sub

Private Sub mnuHelpLegal_Click()
    On Error Resume Next
    Call Me.Hide
    Call frmGutenberg.mnuHelpLegal_Click
End Sub

Private Sub mnuHelpManual_Click()
    On Error Resume Next
    Call Me.Hide
    Call frmGutenberg.mnuHelpManual_Click
End Sub

Private Sub mnuOptionsFont_Click()
    On Error Resume Next
    Dim sel As Long
    Dim selL As Long
    cdlgFont.Flags = cdlCFEffects
    cdlgFont.Color = rtfMain.SelColor
    On Error GoTo cancelError:
    cdlgFont.ShowFont
    On Error Resume Next
    sel = rtfMain.SelStart
    selL = rtfMain.SelLength
    rtfMain.SelStart = 0
    rtfMain.SelLength = Len(rtfMain.text)
    rtfMain.Font.bold = cdlgFont.FontBold
    rtfMain.Font.name = cdlgFont.FontName
    rtfMain.Font.size = cdlgFont.FontSize
    rtfMain.Font.Underline = cdlgFont.FontUnderline
    rtfMain.SelStart = sel
    rtfMain.SelLength = selL
cancelError:
End Sub

Private Sub mnuViewWhat_Click(index As Integer)
    On Error Resume Next
    If index <> 1 Then
        'back to main form
        Call Me.Hide
        Call frmGutenberg.mnuViewWhat_Click(index)
    End If
End Sub

Private Sub rtfMain_KeyDown(KeyCode As Integer, Shift As Integer)
'    On Error Resume Next
'    Dim start As Long
'    Dim what As String
'
'    If KeyCode = vbKeyF And (Shift And vbCtrlMask) > 0 Then
'        what = InputBox("Find what:", "Find", mFind)
'        If Len(what) > 0 Then
'            mFind = what
'            start = frmBook.rtfMain.SelStart + 2
'            If start > Len(rtfMain.text) Then start = Len(rtfMain.text)
'            start = InStr(start, rtfMain.text, what, vbTextCompare)
'            If start = 0 Then
'                'not found, start from beginning again
'                start = InStr(1, rtfMain.text, what, vbTextCompare)
'                If start = 0 Or start = rtfMain.SelStart Then
'                    'not found at all
'                    Call Beep
'                Else
'                    rtfMain.SelStart = start
'                    rtfMain.SelLength = Len(mFind)
'                End If
'            Else
'                'found!
'                rtfMain.SelStart = start - 1
'                rtfMain.SelLength = Len(mFind)
'            End If
'            Call rtfMain.SetFocus
'        End If
'    End If
End Sub

Private Sub rtfMain_LostFocus()
    On Error Resume Next
    If Len(rtfMain.text) > 100 Then
        Call SaveCurrentCursorPosition
    End If
End Sub

Private Sub SaveCurrentCursorPosition()
    On Error Resume Next
    'saves the cursor position in the current book, if any
    If gLoadedIndex > -1 And gLoadedIndex < gBooks.documentElement.selectNodes("book").length Then
        gBooks.documentElement.selectNodes("book").Item(gLoadedIndex).selectSingleNode("cursorPos").text = frmBook.rtfMain.SelStart
    End If
End Sub

