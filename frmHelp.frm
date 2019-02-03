VERSION 5.00
Begin VB.Form frmHelp 
   Caption         =   "Accessible Gutenberg Help"
   ClientHeight    =   5175
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9930
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmHelp.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5175
   ScaleWidth      =   9930
   StartUpPosition =   3  'Windows Default
   Tag             =   "frmHelp"
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "OK"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1680
      TabIndex        =   1
      Top             =   2760
      Width           =   1215
   End
   Begin VB.TextBox txtHelp 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2655
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   0
      Width           =   4695
   End
End
Attribute VB_Name = "frmHelp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mManuallySet As Boolean

Private Sub cmdOK_Click()
    On Error Resume Next
    Call Me.Hide
    Call frmGutenberg.SetFocus
End Sub

Private Sub Form_Activate()
    On Error Resume Next
    Call txtHelp.SetFocus
End Sub

Private Sub Form_Initialize()
    On Error Resume Next
    Call modXPStyle.InitCommonControlsVB
End Sub

Private Sub Form_Load()
    On Error Resume Next
    Call modI18N.ApplyUILanguageToThisForm(Me)
    Call modLargeFonts.ApplySystemSettingsToForm(Me, , True)
    Call modRememberPosition.LoadPosition(Me)
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    If Me.WindowState <> vbMinimized Then
        If Me.Height > cmdOK.Height Then
            cmdOK.Top = Me.ScaleHeight - cmdOK.Height - 90
            txtHelp.Height = Me.ScaleHeight - cmdOK.Height - 270
            txtHelp.Top = 90
        End If
        If Me.Width > cmdOK.Width + 180 Then
            txtHelp.Left = 90
            cmdOK.Left = Me.ScaleWidth / 2 - cmdOK.Width / 2
            txtHelp.Width = Me.ScaleWidth - 180
        End If
    End If
End Sub

Public Sub ShowLegal()
    On Error Resume Next
    txtHelp.text = modI18N.helpTopicText(1)
    Me.caption = modI18N.helpTopicTitle(1)
End Sub

Public Sub ShowHelp()
    On Error Resume Next
    txtHelp.text = modI18N.helpTopicText(0)
    Me.caption = modI18N.helpTopicTitle(0)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    Call modRememberPosition.SavePosition(Me)
End Sub
