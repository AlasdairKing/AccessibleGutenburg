VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Object = "{D0D653FB-B36F-4918-9648-3C495E456DC4}#1.5#0"; "UniBox10.ocx"
Begin VB.Form frmGutenberg 
   Caption         =   "Accessible Gutenberg"
   ClientHeight    =   6630
   ClientLeft      =   225
   ClientTop       =   825
   ClientWidth     =   11880
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   12
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmGutenberg.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   6630
   ScaleWidth      =   11880
   StartUpPosition =   3  'Windows Default
   Begin UniToolbox.UniList lstCategory 
      Height          =   630
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   1575
      _Version        =   65541
      _ExtentX        =   2778
      _ExtentY        =   1111
      _StockProps     =   77
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      IntegralHeight  =   0   'False
      UList           =   ""
   End
   Begin UniToolbox.UniList lstBooks 
      Height          =   630
      Left            =   5160
      TabIndex        =   3
      Top             =   480
      Width           =   1575
      _Version        =   65541
      _ExtentX        =   2778
      _ExtentY        =   1111
      _StockProps     =   77
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      IntegralHeight  =   0   'False
      UList           =   ""
   End
   Begin VB.OptionButton optSort 
      Caption         =   "Sort by &Title"
      Height          =   495
      Index           =   0
      Left            =   5280
      TabIndex        =   6
      Top             =   3360
      Value           =   -1  'True
      Width           =   2295
   End
   Begin VB.OptionButton optSort 
      Caption         =   "Sort by &Author"
      Height          =   495
      Index           =   1
      Left            =   5160
      TabIndex        =   7
      Top             =   3720
      Width           =   2295
   End
   Begin ComctlLib.StatusBar staMain 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   8
      Top             =   6255
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   661
      Style           =   1
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   1
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.TextBox txtDummy 
      Height          =   495
      Left            =   9360
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   4920
      Width           =   1215
   End
   Begin SHDocVwCtl.WebBrowser webBrowser 
      Height          =   1455
      Left            =   8160
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   840
      Width           =   1815
      ExtentX         =   3201
      ExtentY         =   2566
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
   Begin MSComDlg.CommonDialog cdlg 
      Left            =   5760
      Top             =   3120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin VB.Timer tmrBrowser 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   3120
      Top             =   1320
   End
   Begin UniToolbox.UniList lstDownloaded 
      Height          =   630
      Left            =   120
      TabIndex        =   5
      Top             =   3360
      Width           =   1575
      _Version        =   65541
      _ExtentX        =   2778
      _ExtentY        =   1111
      _StockProps     =   77
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      IntegralHeight  =   0   'False
      Sorted          =   -1  'True
      UList           =   "000d000a"
   End
   Begin VB.Label lblMain 
      AutoSize        =   -1  'True
      Caption         =   "Book &Category"
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1800
   End
   Begin VB.Label lblBooks 
      AutoSize        =   -1  'True
      Caption         =   "Book &List"
      Height          =   285
      Left            =   5160
      TabIndex        =   2
      Top             =   120
      Width           =   1125
   End
   Begin VB.Label lblDownloaded 
      AutoSize        =   -1  'True
      Caption         =   "Downloaded &Books"
      Height          =   285
      Left            =   120
      TabIndex        =   4
      Top             =   3000
      Width           =   2325
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileOpen 
         Caption         =   "&Import a book"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuFileSaveas 
         Caption         =   "&Export a book"
         Enabled         =   0   'False
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
      End
      Begin VB.Menu mnuBar 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuRecent 
         Caption         =   ""
         Index           =   0
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuCut 
         Caption         =   "Cu&t"
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
      Begin VB.Menu mnuEditFind 
         Caption         =   "&Find"
         Shortcut        =   ^F
      End
      Begin VB.Menu mnuEditFindnext 
         Caption         =   "Find &next"
         Shortcut        =   {F3}
      End
      Begin VB.Menu mnuSelectall 
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
Attribute VB_Name = "frmGutenberg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Program to let users access the Gutenberg catalog online with no fuss.

Option Explicit

'VERSION HISTORY
'1.0.0
'Sent to Roger to review
'1.0.1
'Added Resume Nexts to all functions
'Fixed ListIndex = 0 sets that might have confused screenreaders.
'Fixed default book
'Fixed authors who had married names from failing to display
'Changed fonts all to Tahoma (Windows XP font)
'Started (but didn't finish!) writing a caching mechanism. Turned out I didn't
'need it: I optimised the webpage processing instead.
'Released, foolishly, Jan 2007.
'1.0.2
'Fixed bug with recent list (no default installed) that made it hang.
'Fixed bug with recent list that allowed duplicate entries.
'1.0.3
'Forgot to update the default book list in the resource file, didn't I? Now
'correctly only shows installed book (Christmas Carol) on installation. Also
'deletes books that are missing from the download list, so it should correct
'previously-broken installs.
'1.0.4
'Fix for non-English language handling.
'1.0.5 1 April 2007
'Change to default langauge handling - see CLanguage.
'1.0.6 25 April 2007
'Wouldn't instantiate InternetExplorer on test machine, so changed reference
'to InternetExplorer to instance of WebBrowser on form.
'1.0.7 15 May 2007
'Updated copyright notice stripping, so it better removes Gutenberg notices.
'Added double-click handling on lists.
'Remembers and displays book currently open on Book tab.
'Fixed bug with shift and page up/down changing tab.
'Remembers the tab you were last in and displays that when opening.
'2.0.0 30 Jan 2008
'Redesigned user interface, added progress bar, message when accessing web.
'2.0.1 23 March 2008
'Fixed i18n bug.
'2.0.2 13 July 2008
'Removed almost all the Gutenberg processing, because it is too unreliable and sometimes
'produces books with no text.
'Fixed Help and Open menu items in Book view.
'Made treeview open when you click right, as well as loading children.
'2.0.3 31 August 2008
'Lets you delete books when list is sorted by author.
'2.0.4  08 Oct 2008
'   Exit while loop error handling in clsDownloader
'2.0.5 23 Mar 2009
'   Fix focus going to web browser with dummy control, hope this fixes freezing reported by Brett Hollis brett.hollis@gmail.com.
'2.0.6 25 May 2009
'   Fixed XP style being misapplied in _Activate, not _Initialize
'2.0.7 8 June 2009
'   Replaced comctl32 TreeView with one from mscomctl because the former would cause the application to crash in JAWS and Narrator.
'2.0.8 13 June 2009
'   Fixed XP Style bug.
'2.0.9 15 June 2009
'   Fixed XP Style bug AGAIN.
'   Fixed books downloading but not being given title.
'   Fixed Books not being opened when selected from Recent.
'3.0.0
'   1 Sep 2009. Changed inaccessible treeview into accessible lists (!
'               Reinstated removal of copyright section.
'               Added Export to allow files to be saved outside the program.
'3.1.0
'  30 Jul 2010. Fixed layout to make sure lists are visible.
'               Took out broken "remember layout" code.
'               Used new MSAA label code to provide MSAA code.

'TO DO
'Check it deletes okay, especially recent list
'Fix removal of copyright notice
'Add some more default books

Private mState As String ' tracks the state of what is being downloaded.
Private mFind As String ' sought by user in Find dialogue
Private mRecent As DOMDocument30 ' holds the most recent files opened by the user.
Private Const MAX_RECENT_LIST As Long = 10 ' how many entries can appear in the recent list in
'   the file menu.

Private Const A As Long = 65
Private Const Z As Long = 90
Private mHREFIndex As Collection 'indexes book node against href so we can get the book

Private mCurrentCategoryIndex As Integer ' The currently-selected category of books/authors

Private mBooklistBooks As Collection ' the books in the lstBooks list

Private Const BOOK_DATA_FILE As String = "\downloadedBooks.xml"
Private Const GUTENBERG_HEADER As String = "*END*THE SMALL PRINT!"
Private Const GUTENBERG_FOOTER As String = "*** END OF THIS PROJECT GUTENBERG EBOOK"

Private Const AUTHOR_FLAG As Long = 1024

Private NO_DOWNLOADED_BOOKS As String

Private Const Document As Integer = 0

Private Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" _
   (ByVal lpBuffer As String, nSize As Long) As Long

Private Declare Function URLDownloadToFile Lib "urlmon" Alias _
    "URLDownloadToFileA" (ByVal pCaller As Long, _
    ByVal szURL As String, _
    ByVal szFileName As String, _
    ByVal dwReserved As Long, _
    ByVal lpfnCB As Long) As Long

Private Sub Form_Initialize()
    On Error Resume Next
    Call modXPStyle.InitCommonControlsVB
End Sub

Private Sub Form_Load()
    On Error Resume Next
    Dim i As Long
    Dim result As String
    
    'create folder for downloading books and storing data
    Call modPath.DetermineSettingsPath("WebbIE", App.title, "1") ' Make sure it keeps reading old store.
    Call modUpdate.CheckForUpdates
    Call modRememberPosition.LoadPosition(Me)
    Call modI18N.ApplyUILanguageToThisForm(Me)
    Call modLargeFonts.ApplySystemSettingsToForm(Me, , True)
    Call modMSAAfyLabels.MSAAfy(Me)
        
    'Set mCurrentCategoryIndex to indicate that we don't have a currently-selected category.
    mCurrentCategoryIndex = -1
    'Set webBrowser = New InternetExplorer
    Call Load(frmBook)
    NO_DOWNLOADED_BOOKS = modI18N.GetText("No books downloaded yet.")
    Call lstBooks.AddItem(GetText("Select a category from the list and press return."))
    'stop user tabbing to IE object
    webBrowser.TabStop = False
    webBrowser.Silent = True
    'no book loaded yet
    gLoadedIndex = -1
    'setup books file
    Call SetupBooksFile
    'setup recent file
    Call SetupRecent
    'setup directory
    Call SetupDirectory
    'load downloaded record
    Call LoadBooks
    'prepare to get book index from Gutenberg
    Set mHREFIndex = New Collection
    Call Me.Show
    'Work out book sort
    result = modPath.GetSettingIni(App.EXEName, "Downloaded", "Sort", "Title")
    If result = "Title" Then
        optSort(0).value = True
        Call optSort_Click(0)
    Else
        optSort(1).value = True
        Call optSort_Click(1)
    End If
    'display current book, if any
    result = modShared.SharedReadIniFile(modPath.settingsPath & "\" & App.title & ".ini", "State", "CurrentBookIndex")
    'Default to first book if not found but a book exists:
    If lstDownloaded.ListCount > 0 And Len(result) = 0 Then result = "0"
    If Len(result) > 0 Then
        lstDownloaded.ListIndex = CInt(result)
        If lstDownloaded.ListIndex > -1 Then
            gLoadedIndex = lstDownloaded.ItemData(lstDownloaded.ListIndex)
            'Call LoadSelectedBook
        End If
    End If
    'restore tab state
    Call Home
    Call Load(frmProgress)
End Sub

Private Sub Home()
    On Error Resume Next
    mState = "home"
    Call lstCategory.SetFocus
    If lstCategory.ListIndex = -1 Then lstCategory.ListIndex = 0
    staMain.SimpleText = Empty
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    Dim i As Long
    If Me.WindowState <> vbMinimized Then
        If Me.ScaleHeight > 2000 Then
            lblMain.Top = GAP
            lstCategory.Top = lblMain.Top + lblMain.Height + GAP / 2
            lstCategory.Height = Me.ScaleHeight / 2 - lblMain.Height
            lblBooks.Top = lblMain.Top
            lstBooks.Top = lstCategory.Top
            lstBooks.Height = lstCategory.Height
            lblDownloaded.Top = lstCategory.Top + lstCategory.Height + GAP
            lstDownloaded.Top = lblDownloaded.Top + lblDownloaded.Height + GAP
            optSort(0).Top = Me.ScaleHeight - optSort(0).Height - staMain.Height
            optSort(1).Top = optSort(0).Top
            lstDownloaded.Height = optSort(0).Top - lstDownloaded.Top
        End If
        If Me.ScaleWidth > 2000 Then
            lblMain.Left = GAP
            lstCategory.Left = GAP
            lstCategory.Width = (Me.ScaleWidth - GAP * 3) / 2
            lstBooks.Width = lstCategory.Width
            lstBooks.Left = lstCategory.Left + lstCategory.Width + GAP
            lblBooks.Left = lstBooks.Left
            lstDownloaded.Left = GAP
            lstDownloaded.Width = Me.ScaleWidth - lstDownloaded.Left
            optSort(0).Left = GAP
            optSort(1).Left = optSort(0).Left + optSort(0).Width + GAP
        End If
    End If
    txtDummy.Left = Me.ScaleWidth + 500
    webBrowser.Left = Me.ScaleWidth + 100
    Call modMSAAfyLabels.Refresh(Me)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    'save downloaded books
    Call gBooks.Save(gBooks.url)
    'save sort type
    If optSort(0).value Then
        Call modPath.SaveSettingIni(App.EXEName, "Downloaded", "Sort", "Title")
    Else
        Call modPath.SaveSettingIni(App.EXEName, "Downloaded", "Sort", "Author")
    End If
    'save current book index
    Call WritePrivateProfileString("State", "CurrentBookIndex", CStr(lstDownloaded.ItemData(lstDownloaded.ListIndex)), modPath.settingsPath & "\" & App.title & ".ini")
    'save position
    Call modRememberPosition.SavePosition(Me)
    'unload forms
    Call Unload(frmHelp)
    Call Unload(frmBook)
    Call Unload(frmProgress)
    Call Unload(frmAccessing)
End Sub

Private Sub lstBooks_DblClick()
    On Error Resume Next
    Call lstBooks_KeyPress(vbKeyReturn)
End Sub

Private Sub lstBooks_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    If KeyCode = vbKeyUp And lstBooks.ListIndex = 0 Then
        Call Beep
    ElseIf KeyCode = vbKeyDown And lstBooks.ListIndex = lstBooks.ListCount - 1 Then
        Call Beep
    End If
End Sub

Private Sub lstBooks_KeyPress(KeyAscii As Integer)
    On Error Resume Next
    Dim href As String
    
    If KeyAscii = vbKeyReturn And lstBooks.ListCount > 0 Then
        'got a book to download. Have we downloaded this already?
        href = mHREFIndex.Item(lstBooks.ListIndex + 1)
        gLoadingID = Right(href, Len(href) - InStrRev(href, "/"))
        If gBooks.documentElement.selectNodes("book[@id=""" & gLoadingID & """]").length > 0 Then
            'we've already downloaded this
            Call LoadSelectedBook(gLoadingID)
        Else
            'download
            mState = "book"
            Call frmAccessing.Show
            Call webBrowser.Navigate2(mHREFIndex.Item(lstBooks.ListIndex + 1))
            'MsgBox "HREF:" & mHREFIndex.item(tvwMain.SelectedItem.key)
        End If
    End If
End Sub

Private Sub lstBooks_KeyUp(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    If KeyCode = vbKeyEscape Then
        Call lstCategory.SetFocus
    End If
End Sub

Private Sub lstDownloaded_DblClick()
    On Error Resume Next
    Call lstDownloaded_KeyPress(vbKeyReturn)
End Sub

Private Sub lstDownloaded_GotFocus()
    On Error Resume Next
    If lstDownloaded.ListIndex = -1 Then lstDownloaded.ListIndex = 0
End Sub

Private Sub lstDownloaded_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    Dim result As Long
    Dim old As Long
    Dim index As Long
    
    If KeyCode = vbKeyDelete Then
        'delete downloaded book
        If lstDownloaded.List(0) = NO_DOWNLOADED_BOOKS Then
            'no books to delete
            Call Beep
        Else
            'delete book
            result = MsgBox(GetText("Do you want to delete") & " " & lstDownloaded.List(lstDownloaded.ListIndex) & " " & GetText("from your list of downloaded books?"), vbYesNoCancel, App.title)
            If result = vbYes Then
                frmBook.rtfMain.text = Empty
                gLoadedIndex = -1
                index = lstDownloaded.ItemData(lstDownloaded.ListIndex)
                Call RemoveRecent(gBooks.documentElement.selectNodes("book").Item(index).selectSingleNode("filename").text)
                Call gBooks.documentElement.removeChild(gBooks.documentElement.selectNodes("book").Item(index))
                Call gBooks.Save(gBooks.url)
                Call LoadBooks
            End If
        End If
    End If
End Sub

Private Sub lstDownloaded_KeyPress(KeyAscii As Integer)
    On Error Resume Next
    If KeyAscii = vbKeyReturn Then
        'load downloaded book
        If lstDownloaded.ListCount = 1 And NO_DOWNLOADED_BOOKS = lstDownloaded.List(0) Then
            'no books!
            Call Beep
        Else
            'downloaded, open
            gLoadedIndex = lstDownloaded.ItemData(lstDownloaded.ListIndex)
            Call LoadSelectedBook
        End If
    End If
End Sub

Private Sub LoadSelectedBook(Optional id As String)
    On Error Resume Next
    Dim Path As String
    Dim cursor As Long
    Dim index As Long
    Dim book As IXMLDOMNode
    Dim name As String
    Dim filename As String
    
    If id = Empty Then
        index = lstDownloaded.ItemData(lstDownloaded.ListIndex)
        If index >= 0 Then
            Set book = gBooks.documentElement.selectNodes("book").Item(index)
        End If
    Else
        Set book = gBooks.documentElement.selectSingleNode("book[@id='" & id & "']")
    End If
    If Not (book Is Nothing) Then
        If frmBook.rtfMain.filename <> "" Then Call frmBook.rtfMain.SaveFile(frmBook.rtfMain.filename)
        Call frmBook.rtfMain.LoadFile(modPath.nonRoamingSettingsPath & "\" & book.selectSingleNode("filename").text)
        Call AddRecent(book.selectSingleNode("title").text, book.selectSingleNode("filename").text)
        cursor = book.selectSingleNode("cursorPos").text
        If cursor > 0 Then
            frmBook.rtfMain.SelStart = cursor
        End If
        frmBook.caption = GetText("Book") & " - " & book.selectSingleNode("title").text
        Call frmBook.Show(vbModal, Me)
    End If
End Sub

Private Sub lstCategory_DblClick()
    On Error Resume Next
    Call lstCategory_KeyPress(vbKeyReturn)
End Sub

Private Sub lstCategory_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    If KeyCode = vbKeyUp And lstCategory.ListIndex = 0 Then
        Call Beep
    ElseIf KeyCode = vbKeyDown And lstCategory.ListIndex = lstCategory.ListCount - 1 Then
        Call Beep
    End If
End Sub

Private Sub lstCategory_KeyPress(KeyAscii As Integer)
    On Error Resume Next
    Dim href As String
    Dim getting As String
    
    'This is where the main downloading activity is initiated.
    If KeyAscii = vbKeyReturn And lstCategory.ListIndex > -1 Then
        KeyAscii = 0
        If mCurrentCategoryIndex = lstCategory.ItemData(lstCategory.ListIndex) Then
            'already got this category, just change focus.
            Call lstBooks.SetFocus
        Else
            mCurrentCategoryIndex = lstCategory.ItemData(lstCategory.ListIndex)
            If lstCategory.ItemData(lstCategory.ListIndex) > AUTHOR_FLAG Then
                'one of the author categories
                staMain.SimpleText = GetText("Getting list of authors...")
                'record which index letter we need.
                getting = LCase(Left(lstCategory.text, 1))
                mState = "authorList"
                'load from Gutenberg: see DocumentComplete for processing
                Call frmAccessing.Show
                Call webBrowser.navigate("http://www.gutenberg.org/browse/authors/" & getting)
            Else
                'One of the title categories
                'one of the title nodes
                staMain.SimpleText = GetText("Getting list of titles...")
                'record which index letter we need.
                getting = LCase(Left(lstCategory.text, 1))
                mState = "titleList"
                'load from Gutenberg: see DocumentComplete for processing
                Call frmAccessing.Show
                Call webBrowser.navigate("http://www.gutenberg.org/browse/titles/" & getting)
            End If
        End If
        KeyAscii = 0
    End If
End Sub

Public Sub mnuFileOpen_Click()
    On Error Resume Next
    Dim name As String
    Dim author As String
    Dim s As String
    Dim fso As New Scripting.FileSystemObject
    Dim entry As Integer
    
    cdlg.Flags = cdlOFNFileMustExist + cdlOFNPathMustExist
    cdlg.DefaultExt = "txt"
    cdlg.cancelError = True
    cdlg.DialogTitle = "Open document"
    cdlg.Filter = "Documents|*.txt;*.doc;*.rtf;*.text"
    On Error GoTo cancelError:
    cdlg.ShowOpen
    On Error Resume Next
    If cdlg.filename <> "" Then
        name = InputBox("Name of document:")
        author = InputBox("Author:")
        If Len(name) = 0 Then
            name = cdlg.filename
            If InStr(1, name, "\") > 0 Then
                name = Right(name, Len(name) - InStrRev(name, "\", Len(name)))
            End If
            If InStr(1, name, ".") > 0 Then
                name = Left(name, Len(name) - InStr(1, name, ".") - 1)
            End If
        End If
        If Len(author) = 0 Then
            s = String(256, Chr(0))
            Call GetUserName(s, Len(s))
            author = Left(s, InStr(1, s, Chr(0)) - 1)
        End If
        If frmBook.rtfMain.filename <> "" Then Call frmBook.rtfMain.SaveFile(frmBook.rtfMain.filename)
        Call frmBook.rtfMain.LoadFile(cdlg.filename)
        'Trim path to just filename
        entry = AddToBooks(name, fso.GetFile(cdlg.filename).name, author, Globals.GetUniqueKey)
        Call AddRecent(name, cdlg.filename)
    End If
cancelError:
End Sub

Public Sub mnuViewWhat_Click(index As Integer)
    On Error Resume Next
    Select Case index
        Case 0 ' catalogue
            Call lstCategory.SetFocus
        Case 1 ' book
            Call LoadSelectedBook
        Case 2 ' downloaded book list
            Call lstDownloaded.SetFocus
    End Select
End Sub

Private Sub webBrowser_DocumentComplete(ByVal pDisp As Object, url As Variant)
    On Error Resume Next
    Dim elementIterator As IHTMLElement
    Dim filename As String
    Dim Path As String
    Dim slash As Long
    Dim notice As Long
    Dim text As String
    Dim href As String
    Dim book As clsBook

'    Debug.Print "url:" & url
    If url = "http:///" Or url = "" Then
    ElseIf webBrowser.readyState = READYSTATE_COMPLETE Then
        Select Case mState
            Case "book"
                'We're trying to download a particular book.
                'get book link and download it
                For Each elementIterator In pDisp.Document.getElementsByTagName("A")
                    filename = elementIterator.getAttribute("href")
                    Path = elementIterator.getAttribute("href")
                    If InStr(1, filename, ".txt", vbTextCompare) > 0 Then
                        'get the filename by parsing href
                        slash = InStrRev(filename, "/", Len(filename))
                        filename = Right(filename, Len(filename) - slash)
                        'filename = Left(filename, InStrRev(filename, "/", Len(filename)))
                        'work out title/author by looking it up in mBooklistBooks
                        Set book = mBooklistBooks.Item(lstBooks.ListIndex + 1)
                        Call frmAccessing.Hide
                        Call frmProgress.SetDownload(Path, modPath.nonRoamingSettingsPath & "\" & filename, book.title)
                        Call frmProgress.Show(vbModal, Me)
                        Exit For
                    End If
                Next elementIterator
            Case Else
                'We've just downloaded a category of books from the website.
                Call DisplayCatalogue(webBrowser.Document)
                Call frmAccessing.Hide
        End Select
    End If
End Sub

Public Sub BookDownloaded(filename As String, url As String)
    On Error Resume Next
    Dim text As String
    Dim author As String
    Dim notice As String
    Dim n As Node
    Dim book As clsBook
    Dim fso As Scripting.FileSystemObject
    Dim ts As Scripting.TextStream
    Dim tempFilename As String
    Dim counter As Long
    Dim copyrightLine As Long
    Dim totalLines As Long
    
    Set book = mBooklistBooks.Item(lstBooks.ListIndex + 1)
    Set fso = New Scripting.FileSystemObject
    'Add to downloaded books
    gLoadedIndex = AddToBooks(book.title, fso.GetFile(filename).name, book.author, gLoadingID)
    'Save existing book
    If frmBook.rtfMain.filename <> "" Then Call frmBook.rtfMain.SaveFile(frmBook.rtfMain.filename)
    'Load new book
    Call frmBook.rtfMain.LoadFile(filename)
    Call AddRecent(book.title, filename)
    'No longer try to strip out Gutenberg stuff: it's too unreliable, sometimes we lose the whole book.
    'The user will just have to find the start themselves.
    'Still process the newlines, though.
'    'trim out the gutenberg header stuff
    tempFilename = fso.GetSpecialFolder(TemporaryFolder) & "\" & fso.GetTempName & ".txt"
    Set ts = fso.CreateTextFile(tempFilename, True, True)
    Call ts.write(frmBook.rtfMain.text)
    Call ts.Close
    DoEvents
    Set ts = fso.OpenTextFile(tempFilename, ForReading, False, TristateTrue)
    counter = 0
    totalLines = 0
    While Not ts.AtEndOfStream
        text = ts.ReadLine
        totalLines = totalLines + 1
        'Look for end of copyright notice
        If copyrightLine > 0 Then
            'Go with the first notice end we find, in case they are duplicated at the bottom.
        ElseIf InStr(1, text, "*** START OF THE PROJECT", vbTextCompare) > 0 Then
            'Found copyright line
            copyrightLine = totalLines
        ElseIf InStr(1, text, "*** START OF THIS PROJECT", vbTextCompare) > 0 Then
            'Found copyright line
            copyrightLine = totalLines
        ElseIf InStr(1, text, "*END THE SMALL PRINT!", vbTextCompare) > 0 Then
            'Found copyright line
            copyrightLine = totalLines
        End If
    Wend
    Call ts.Close
    DoEvents
    'OK, so totalLines has the number of lines, and copyrightLine has the line of the end of the top copyright
    'section, if any found.
    If copyrightLine > 0 And (copyrightLine < (totalLines * 0.25)) Then
        'If the copyright notice does not come in the top 25% of the text, assume we've incorrectly found
        'the BOTTOM copyright notice, and do nothing.
        'Now chop the file.
        Set ts = fso.OpenTextFile(tempFilename, ForReading, False, TristateTrue)
        counter = 0
        While Not ts.AtEndOfStream
            If counter <= copyrightLine Then
                'Still in copyright section
                'Read line and throw it away.
                Debug.Print "Dumping: " & ts.ReadLine
            Else
                text = text & ts.ReadLine & vbNewLine
            End If
            counter = counter + 1
        Wend
        Call ts.Close
        DoEvents
        text = book.title & " - " & book.author & vbNewLine & vbNewLine & text
    Else
        text = frmBook.rtfMain.text
    End If
    Call fso.DeleteFile(tempFilename)
'    'header:
'    notice = InStr(1, text, GUTENBERG_HEADER, vbTextCompare)
'    If notice > 0 Then
'        notice = InStr(notice, text, vbNewLine, vbTextCompare)
'        text = Right(text, Len(text) - notice + 1)
'    Else
'        'hmm, didn't find header. Maybe different format.
'        notice = InStr(1, text, vbNewLine & vbNewLine & vbNewLine & vbNewLine & vbNewLine & vbNewLine, vbBinaryCompare)
'        If notice > 0 Then
'            'six newlines: that's probably the header
'            notice = notice + Len(vbNewLine) * 6
'            text = Right(text, Len(text) - notice + 1)
'        End If
'    End If
'    text = Trim(text)
'    While Left(text, Len(vbNewLine)) = vbNewLine And Len(text) > 0
'        text = Right(text, Len(text) - Len(vbNewLine))
'    Wend
'    'footer:
'    notice = InStr(1, text, GUTENBERG_FOOTER, vbTextCompare)
'    If notice > 0 Then
'        text = Left(text, notice)
'    End If
'    text = Trim(text)
'    While Left(text, Len(vbNewLine)) = vbNewLine
'        text = Right(text, Len(text) - Len(vbNewLine))
'    Wend
    'take out the extraneous newlines
    text = Replace(text, vbNewLine & vbNewLine, "jkjkjkjkjjkjkjk")
    text = Replace(text, vbNewLine, " ")
    text = Replace(text, "  ", " ")
    text = Replace(text, "jkjkjkjkjjkjkjk", vbNewLine & vbNewLine)
    frmBook.rtfMain.text = text
    Call frmBook.rtfMain.SaveFile(filename)
    staMain.SimpleText = Empty
    Call LoadSelectedBook(gLoadingID)
End Sub

Private Sub DisplayCatalogue(doc As IHTMLDocument)
'Load entries from the page and put in main list.
    On Error Resume Next
    Dim elementIterator As IHTMLElement
    Dim listcat As String
    Dim publications As IHTMLElement
    Dim book As IHTMLElement
    Dim newKey As String
    Dim newKey2 As String
    Dim childA As IHTMLElement
    Dim foundBook As Boolean
    Dim author As IHTMLElement
    Dim authorName As String
    Dim title As String
    Dim newBook As clsBook
    
    staMain.SimpleText = GetText("Loading catalogue - please wait")
    'Call modShared.LockWindowUpdate(tvwMain.hwnd)
    Call lstBooks.Clear
    Set mBooklistBooks = New Collection
    For Each elementIterator In doc.getElementsByTagName("H2")
        authorName = (Replace(elementIterator.innerText, vbNewLine, " "))
        Set publications = elementIterator.nextSibling
        foundBook = False
        newKey = Globals.GetUniqueKey
        If mState = "authorList" Then
            For Each book In publications.children
                newKey2 = Globals.GetUniqueKey
                If book.tagName = "LI" Then
                    If book.className = "pgdbetext" Then
                        'Got an etext to add to this author.
                        title = Replace(book.innerText, vbNewLine, " ")
                        Call lstBooks.AddItem(authorName & " - " & title)
                        For Each childA In book.children
                            If childA.tagName = "A" Then
                                foundBook = True
                                Call mHREFIndex.Add(childA.getAttribute("href"), newKey2)
                                Exit For
                            End If
                        Next childA
                        Set newBook = New clsBook
                        newBook.author = authorName
                        newBook.title = title
                        Call mBooklistBooks.Add(newBook)
                    End If
                End If
            Next book
        ElseIf mState = "titleList" Then
            Call mHREFIndex.Add(elementIterator.children(0).getAttribute("href"), newKey)
            Set author = elementIterator.nextSibling
            authorName = author.innerText
            title = Replace(elementIterator.innerText, vbNewLine, " ")
            Set newBook = New clsBook
            newBook.title = title
            newBook.author = authorName
            Call mBooklistBooks.Add(newBook)
            Call lstBooks.AddItem(title & " - " & authorName)
            foundBook = True
        End If
    Next elementIterator
    staMain.SimpleText = Empty
    lstBooks.ListIndex = 0
    Call lstBooks.SetFocus
End Sub

Private Sub webBrowser_StatusTextChange(ByVal text As String)
    On Error Resume Next
    staMain.SimpleText = text
End Sub

Private Sub mnuCopy_Click()
    On Error Resume Next
    Call SendKeys("^C")
End Sub

Private Sub mnuCut_Click()
    On Error Resume Next
    Call SendKeys("^X")
End Sub

Private Sub mnuEditFind_Click()
    On Error Resume Next
    Dim newSearch As String
    
    newSearch = InputBox("Find what:", "Find", mFind)
    If Len(newSearch) > 0 Then
        mFind = newSearch
        Call Find(mFind)
    End If
End Sub

Private Sub Find(what As String)
    On Error Resume Next
    Dim searchList As ListBox
    Dim start As Long
    Dim i As Long
    Dim found As Boolean

    If Len(what) > 0 Then
        mFind = what
        'Debug.Print "Name:" & Me.ActiveControl.Name
        Select Case Me.ActiveControl.name
            Case lstBooks.name
                Call FindInList(lstBooks, what)
            Case lstDownloaded.name
                Call FindInList(lstDownloaded, what)
        End Select
    End If
End Sub

Private Sub FindInTree(t As TreeView, what As String)
    On Error Resume Next
    Dim i As Long
    Dim keepLooking As Boolean
    Dim found As Node
    
    i = t.SelectedItem.index + 1
    keepLooking = (i <= t.Nodes.Count) ' don't search current item.
    While keepLooking
        If InStr(1, t.Nodes.Item(i), what, vbTextCompare) > 0 Then
            keepLooking = False
            Set found = t.Nodes.Item(i)
        Else
            If i = t.Nodes.Count Then
                keepLooking = False
            Else
                i = i + 1
            End If
        End If
    Wend
    If found Is Nothing Then
        'didn't find, search again from the top.
        keepLooking = True
        i = 1
        While keepLooking
            If InStr(1, t.Nodes.Item(i), what, vbTextCompare) > 0 Then
                keepLooking = False
                Set found = t.Nodes.Item(i)
            Else
                If i = t.SelectedItem.index Then
                    keepLooking = False
                Else
                    i = i + 1
                End If
            End If
        Wend
    End If
    If found Is Nothing Then
        'failed to find
    Else
        'found matching text
        Set t.SelectedItem = found
    End If
End Sub

Private Sub FindInList(l As UniList, what As String)
    On Error Resume Next
    Dim found As Boolean
    Dim start As Long
    Dim i As Long
    
    start = l.ListIndex
    For i = l.ListIndex + 1 To l.ListCount - 1
        If InStr(1, l.List(i), what, vbTextCompare) > 0 Then
            l.ListIndex = i
            found = True
            Exit For
        End If
    Next i
    If Not found Then
        For i = 0 To l.ListIndex - 1
            If InStr(1, l.List(i), what, vbTextCompare) > 0 Then
                l.ListIndex = i
                found = True
                Exit For
            End If
        Next i
    End If
    If found Then
        'okay!
    Else
        Call Beep
    End If
    Call l.SetFocus
End Sub


Private Sub mnuEditFindnext_Click()
    On Error Resume Next
    Call Find(mFind)
End Sub

Private Sub mnuFileExit_Click()
    On Error Resume Next
    Call Unload(Me)
End Sub

Public Sub mnuHelpAbout_Click()
    On Error Resume Next
    Call MsgBox(App.title & " " & App.Major & "." & App.Minor & "." & App.Revision & vbNewLine & "WebbIE Suite " & modVersion.GetPackageVersion & vbNewLine & "Alasdair King http://www.webbie.org.uk", vbInformation, App.title)
End Sub


Public Sub mnuHelpLegal_Click()
    On Error Resume Next
    Call frmHelp.ShowLegal
    Call frmHelp.Show(vbModeless, Me)
End Sub

Public Sub mnuHelpManual_Click()
    On Error Resume Next
    Call frmHelp.ShowHelp
    Call frmHelp.Show(vbModeless, Me)
End Sub

Private Sub mnuPaste_Click()
    On Error Resume Next
    Call SendKeys("^V")
End Sub

Private Sub mnuRecent_Click(index As Integer)
    On Error Resume Next
    Dim Path As String
    Dim caption As String
    'load downloaded book
    If frmBook.rtfMain.filename <> "" Then Call frmBook.rtfMain.SaveFile(frmBook.rtfMain.filename)
    Path = modPath.nonRoamingSettingsPath & "\" & mRecent.documentElement.selectNodes("recent").Item(index).selectSingleNode("filename").text
    Call frmBook.rtfMain.LoadFile(Path)
    caption = GetText("frmBook.Caption")
    If caption <> "frmBook.Caption" Then
        frmBook.caption = caption & " - " & mRecent.documentElement.selectNodes("recent").Item(index).selectSingleNode("name").text
    End If
    Call frmBook.Show(vbModal, Me)
End Sub

Private Sub mnuSelectall_Click()
    On Error Resume Next
    frmBook.rtfMain.SelStart = 1
    frmBook.rtfMain.SelLength = Len(frmBook.rtfMain.text)
End Sub

Private Sub webBrowser_GotFocus()
    On Error Resume Next
    'Need to put in dummy controls in case of known Microsoft bug breaking MSAA.
    Call Me.txtDummy.SetFocus
    'Call Me.tvwMain.SetFocus
End Sub

Private Sub optSort_Click(index As Integer)
    On Error Resume Next
    Call LoadBooks
End Sub

Private Sub DisplayDownloaded()
    On Error Resume Next
    Call lstDownloaded.ZOrder
    Call lstDownloaded.SetFocus
End Sub

Private Sub LoadBooks()
'load the reference file describing the books already downloaded
'and displays them in lstDownloaded
    On Error Resume Next
    Dim bookIterator As IXMLDOMNode
    Dim i As Long
    Dim counter
    Dim fso As FileSystemObject
    
    Set gBooks = New DOMDocument30
    gBooks.async = False
    Call gBooks.Load(modPath.nonRoamingSettingsPath & BOOK_DATA_FILE)
    Call lstDownloaded.Clear
    Set fso = New FileSystemObject
    counter = 0
    For Each bookIterator In gBooks.documentElement.selectNodes("book")
        'Check this book is actually there: need to do this because I cocked
        'up the default file once, so I need to edit the book list if one
        'of them is missing.
        If fso.FileExists(modPath.nonRoamingSettingsPath & "\" & bookIterator.selectSingleNode("filename").text) Then
            If Me.optSort(0).value Then
                'Sort by title
                Call lstDownloaded.AddItem(bookIterator.selectSingleNode("title").text & " - " & bookIterator.selectSingleNode("author").text)
            Else
                Call lstDownloaded.AddItem(bookIterator.selectSingleNode("author").text & " - " & bookIterator.selectSingleNode("title").text)
            End If
            'Store a reference to which book this is.
            lstDownloaded.ItemData(lstDownloaded.NewIndex) = counter
            counter = counter + 1
        Else
            'Book isn't actually downloaded: need to delete from book list
            Call gBooks.documentElement.removeChild(bookIterator)
        End If
    Next bookIterator
    If lstDownloaded.ListCount = 0 Then
        Call lstDownloaded.AddItem(NO_DOWNLOADED_BOOKS)
    End If
    If lstDownloaded.ListCount = 0 Then
        Call lstDownloaded.AddItem(NO_DOWNLOADED_BOOKS)
    End If
    'Don't set ListIndex unless you mean a screenreader to change focus
    'to that list
    'lstDownloaded.ListIndex = 0
End Sub

Private Sub SetupBooksFile()
    On Error Resume Next
    Dim fso As FileSystemObject
    Dim ts As TextStream
    Dim b() As Byte
    Dim text As String
    Dim testDoc As DOMDocument30
    
    Set fso = New FileSystemObject
    If fso.FileExists(modPath.nonRoamingSettingsPath & BOOK_DATA_FILE) Then
        'okay! Does it parse? This can happen when you've added some Chinese text (!)
    Else
        'no books file: copy default book list, which defaults to English. (3.0)
        Call fso.CopyFile(App.Path & BOOK_DATA_FILE, modPath.nonRoamingSettingsPath & BOOK_DATA_FILE, True)
        'hard code this stuff for the time being!
        b() = VB.LoadResData("EN1", "BOOKDATAFILES")
        text = StrConv(b(), vbUnicode)
        Set ts = fso.OpenTextFile(modPath.nonRoamingSettingsPath & "\46-8.txt", ForWriting, True)
        Call ts.write(text)
        Call ts.Close
'       DEV: This makes 6MB of installer, so pulled them out. Only do Christmas Carol.
'        b() = VB.LoadResData("EN2", "BOOKDATAFILES")
'        text = StrConv(b(), vbUnicode)
'        Set ts = fso.OpenTextFile(modPath.nonRoamingSettingsPath & "\alice30.txt", ForWriting, True)
'        Call ts.write(text)
'        Call ts.Close
'
'        b() = VB.LoadResData("EN3", "BOOKDATAFILES")
'        text = StrConv(b(), vbUnicode)
'        Set ts = fso.OpenTextFile(modPath.nonRoamingSettingsPath & "\dracu13.txt", ForWriting, True)
'        Call ts.write(text)
'        Call ts.Close
'
'        b() = VB.LoadResData("EN4", "BOOKDATAFILES")
'        text = StrConv(b(), vbUnicode)
'        Set ts = fso.OpenTextFile(modPath.nonRoamingSettingsPath & "\frank15.txt", ForWriting, True)
'        Call ts.write(text)
'        Call ts.Close
'
'        b() = VB.LoadResData("EN5", "BOOKDATAFILES")
'        text = StrConv(b(), vbUnicode)
'        Set ts = fso.OpenTextFile(modPath.nonRoamingSettingsPath & "\janey11.txt", ForWriting, True)
'        Call ts.write(text)
'        Call ts.Close
'
'        b() = VB.LoadResData("EN6", "BOOKDATAFILES")
'        text = StrConv(b(), vbUnicode)
'        Set ts = fso.OpenTextFile(modPath.nonRoamingSettingsPath & "\lwmen13.txt", ForWriting, True)
'        Call ts.write(text)
'        Call ts.Close
'
'        b() = VB.LoadResData("EN7", "BOOKDATAFILES")
'        text = StrConv(b(), vbUnicode)
'        Set ts = fso.OpenTextFile(modPath.nonRoamingSettingsPath & "\olivr11.txt", ForWriting, True)
'        Call ts.write(text)
'        Call ts.Close
'
'        b() = VB.LoadResData("EN8", "BOOKDATAFILES")
'        text = StrConv(b(), vbUnicode)
'        Set ts = fso.OpenTextFile(modPath.nonRoamingSettingsPath & "\pandp12.txt", ForWriting, True)
'        Call ts.write(text)
'        Call ts.Close
'
'        b() = VB.LoadResData("EN9", "BOOKDATAFILES")
'        text = StrConv(b(), vbUnicode)
'        Set ts = fso.OpenTextFile(modPath.nonRoamingSettingsPath & "\tarzn10.txt", ForWriting, True)
'        Call ts.write(text)
'        Call ts.Close
'
    End If
End Sub

Private Function AddToBooks(title As String, filename As String, author As String, id As String) As Integer
    On Error Resume Next
    'add a book to the saved book list and return where it was placed (0-based)
    Dim book As IXMLDOMNode
    Dim added As Boolean
    Dim newNode As IXMLDOMNode
    
    
    Set newNode = gBooks.createNode(NODE_ELEMENT, "book", Empty)
    Call newNode.appendChild(gBooks.createNode(NODE_ELEMENT, "title", Empty))
    newNode.selectSingleNode("title").text = title
    Call newNode.appendChild(gBooks.createNode(NODE_ELEMENT, "filename", Empty))
    newNode.selectSingleNode("filename").text = filename
    Call newNode.appendChild(gBooks.createNode(NODE_ELEMENT, "author", Empty))
    newNode.selectSingleNode("author").text = author
    Call newNode.appendChild(gBooks.createNode(NODE_ELEMENT, "cursorPos", Empty))
    newNode.selectSingleNode("cursorPos").text = "1"
    Call newNode.Attributes.setNamedItem(gBooks.createNode(NODE_ATTRIBUTE, "id", Empty))
    newNode.Attributes.getNamedItem("id").text = id
    AddToBooks = 0
    'trim title: appears to gain an extra space at times
    title = Trim(title)
    For Each book In gBooks.documentElement.selectNodes("book")
        'Debug.Print "comparing title=[" & title & "]" & vbNewLine & " with node=[" & book.selectSingleNode("title").text & "]"
        If StrComp(book.selectSingleNode("title").text, title, vbTextCompare) = 0 Then
            'found it already in there!
            Debug.Print "Found already"
            added = True
            Exit For
        ElseIf StrComp(book.selectSingleNode("title").text, title, vbTextCompare) > 0 Then
            added = True
            Debug.Print "Added new"
            Call gBooks.documentElement.insertBefore(newNode, book)
            Exit For
        End If
        AddToBooks = AddToBooks + 1
    Next book
    If Not added Then
        Call gBooks.documentElement.appendChild(newNode)
    End If
    Call gBooks.Save(gBooks.url)
    Call LoadBooks
End Function

Private Sub SetupRecent()
    On Error Resume Next
    Dim fso As FileSystemObject
    Dim b() As Byte
    Dim ts As TextStream
    
    Set fso = New FileSystemObject
    If Not fso.FileExists(modPath.nonRoamingSettingsPath & "\recent.xml") Then
        b() = VB.LoadResData("COMMON", "RECENT")
        Set ts = fso.CreateTextFile(modPath.nonRoamingSettingsPath & "\recent.xml", True)
        Call ts.write(StrConv(b(), vbUnicode))
        Call ts.Close
    End If
    Set mRecent = New DOMDocument30
    mRecent.async = False
    
    Call mRecent.Load(modPath.nonRoamingSettingsPath & "\recent.xml")
    Call DisplayRecent
End Sub

Private Sub DisplayRecent()
    On Error Resume Next
    Dim recentIterator As IXMLDOMNode
    Dim i As Long
    
    'clear recent
    For i = mnuRecent.UBound To 1 Step -1
        Call Unload(mnuRecent(i))
    Next i
    mnuRecent(0).Visible = False
    'load new
    i = 0
    'Debug.Print "mRecent" & vbNewLine & mRecent.xml
    For Each recentIterator In mRecent.documentElement.selectNodes("recent")
        i = i + 1
        If i - 1 > mnuRecent.UBound Then
            Call Load(mnuRecent(i - 1))
        End If
        mnuRecent(i - 1).Visible = True
    Next recentIterator
    i = 0
    For Each recentIterator In mRecent.documentElement.selectNodes("recent")
        'Debug.Print "RX" & recentIterator.xml
        mnuRecent(i).caption = "&" & (i + 1) & " " & recentIterator.selectSingleNode("name").text
        i = i + 1
    Next recentIterator
    mnuBar.Visible = mnuRecent(0).Visible
End Sub

Private Sub AddRecent(name As String, filename As String)
    On Error Resume Next
    Dim newNode As IXMLDOMNode
    Dim nodeIterator As IXMLDOMNode
    Dim needsToBeAdded As Boolean
    
    'first see if we already have the filename
    needsToBeAdded = True
    
    For Each nodeIterator In mRecent.documentElement.selectNodes("recent")
        If nodeIterator.selectSingleNode("filename").text = filename Then
            'we already have this in the recent list: rest easy
            needsToBeAdded = False
            Exit For
        End If
    Next nodeIterator
    'okay, is this something new to add?
    If needsToBeAdded Then
        'yes, it is!
        Set newNode = mRecent.createNode(NODE_ELEMENT, "recent", "")
        Call newNode.appendChild(mRecent.createNode(NODE_ELEMENT, "name", ""))
        newNode.selectSingleNode("name").text = name
        Call newNode.appendChild(mRecent.createNode(NODE_ELEMENT, "filename", ""))
        newNode.selectSingleNode("filename").text = filename
        'add to beginning
        Call mRecent.documentElement.insertBefore(newNode, mRecent.documentElement.selectNodes("recent").Item(0))
        'remove last one if > MAX_RECENT_LIST
        If mRecent.documentElement.selectNodes("recent").length > MAX_RECENT_LIST Then
            Call mRecent.documentElement.removeChild(mRecent.documentElement.selectNodes("recent").Item(5))
        End If
        'save amended file
        Call mRecent.Save(modPath.nonRoamingSettingsPath & "\recent.xml")
        'show list
    End If
    Call DisplayRecent
End Sub

Private Sub RemoveRecent(filename As String)
    On Error Resume Next
'removes an item from the recent list because it's been deleted
    Dim recentIterator As IXMLDOMNode
    
    For Each recentIterator In mRecent.documentElement.selectNodes("recent")
        If recentIterator.selectSingleNode("filename").text = filename Then
            'remove node
            Call mRecent.documentElement.removeChild(recentIterator)
            Call mRecent.Save(mRecent.url)
            Call DisplayRecent
            Exit For
        End If
    Next recentIterator
End Sub

Private Sub SetupDirectory()
    On Error Resume Next
    Dim i As Integer
    Call lstCategory.Clear
    For i = A To Z
        Call lstCategory.AddItem(Chr(i) & vbTab & GetText("Author"))
        lstCategory.ItemData(lstCategory.NewIndex) = i + AUTHOR_FLAG
    Next i
    For i = A To Z
        Call lstCategory.AddItem(Chr(i) & vbTab & GetText("Title"))
        lstCategory.ItemData(lstCategory.NewIndex) = i
    Next i
    
'
'    Dim i As Integer
'    Call tvwMain.Nodes.Clear
'    Call tvwMain.Nodes.Add(, , "home", "Gutenberg Directory")
'    Call tvwMain.Nodes.Add("home", tvwChild, "authors", "Books by Author")
'    Call tvwMain.Nodes.Add("home", tvwChild, "titles", "Books by Title")
'    For i = vbKeyA To vbKeyZ
'        Call tvwMain.Nodes.Add("authors", tvwChild, "authors-" & Chr(i), Chr(i))
'        Call tvwMain.Nodes.Add("titles", tvwChild, "titles-" & Chr(i), Chr(i))
'    Next i
End Sub

Private Sub txtDummy_GotFocus()
    On Error Resume Next
    Call Me.lstCategory.SetFocus
End Sub
