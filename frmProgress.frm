VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form frmProgress 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Book download progress"
   ClientHeight    =   1170
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6915
   ControlBox      =   0   'False
   Icon            =   "frmProgress.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1170
   ScaleWidth      =   6915
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin ComctlLib.ProgressBar prgDownload 
      Height          =   375
      Left            =   90
      TabIndex        =   1
      Top             =   480
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   661
      _Version        =   327682
      Appearance      =   1
   End
   Begin VB.Label lblProgress 
      AutoSize        =   -1  'True
      Caption         =   "&Progress"
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   615
   End
End
Attribute VB_Name = "frmProgress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents mDownloader As clsDownloader
Attribute mDownloader.VB_VarHelpID = -1
Private mURL As String
Private mFilename As String
Private mName As String

Private Sub Form_Activate()
    On Error Resume Next
    Call mDownloader.Download(mURL, mFilename)
End Sub

Private Sub Form_Initialize()
    On Error Resume Next
    Call modXPStyle.InitCommonControlsVB
End Sub

Private Sub Form_Load()
    On Error Resume Next
    Call modI18N.ApplyUILanguageToThisForm(Me)
    Call modLargeFonts.ApplySystemSettingsToForm(Me, , True)
    Call modMSAAfyLabels.MSAAfy(Me)
    Set mDownloader = New clsDownloader
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    lblProgress.Left = GAP
    lblProgress.Top = GAP
    prgDownload.Left = lblProgress.Left
    prgDownload.Top = lblProgress.Top + lblProgress.Height + GAP
    Me.Left = Screen.Width / 2 - Me.Width / 2
    Me.Top = Screen.Height / 2 - Me.Height / 2
    Call Me.Refresh
End Sub

Public Sub SetDownload(url As String, filename As String, Optional name As String)
    On Error Resume Next
    prgDownload.value = 0
    mURL = url
    mFilename = filename
    mName = name
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    Call modRememberPosition.SavePosition(Me)
End Sub

Private Sub mDownloader_Complete(filename As String, url As String)
    On Error Resume Next
    Call Me.Hide
    Call frmGutenberg.BookDownloaded(filename, url)
End Sub

Private Sub mDownloader_Failed(message As String)
    On Error Resume Next
    MsgBox message, vbExclamation
    Call Me.Hide
End Sub

Private Sub mDownloader_Progress(retrieved As Long, total As Long)
    On Error Resume Next
    Dim s As String
    s = Round(retrieved / total, 2) * 100 & "% "
    If Len(mName) > 0 Then
        s = s & modI18N.GetText("of") & " " & mName
    End If
    lblProgress.caption = s
    prgDownload.value = Round(retrieved / total, 2) * 100
End Sub
