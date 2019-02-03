VERSION 5.00
Begin VB.Form frmAccessing 
   BorderStyle     =   0  'None
   Caption         =   "Accessing Gutenberg..."
   ClientHeight    =   3090
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4680
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   12
      Charset         =   204
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.Line linRight 
      X1              =   4680
      X2              =   4680
      Y1              =   0
      Y2              =   360
   End
   Begin VB.Line linLeft 
      BorderWidth     =   2
      X1              =   0
      X2              =   0
      Y1              =   0
      Y2              =   1680
   End
   Begin VB.Line linBottom 
      X1              =   0
      X2              =   4680
      Y1              =   360
      Y2              =   360
   End
   Begin VB.Line linTop 
      X1              =   0
      X2              =   4680
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Label lblAccessing 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Accessing Gutenberg, please wait..."
      Height          =   300
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   4335
   End
End
Attribute VB_Name = "frmAccessing"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Initialize()
    On Error Resume Next
    Call modXPStyle.InitCommonControlsVB
End Sub

Private Sub Form_Load()
    On Error Resume Next
    Call modI18N.ApplyUILanguageToThisForm(Me)
    Call modLargeFonts.ApplySystemSettingsToForm(Me, , True)
    Call modMSAAfyLabels.MSAAfy(Me)
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    lblAccessing.Left = GAP
    lblAccessing.Top = GAP
    Me.Width = lblAccessing.Width + GAP + GAP
    Me.Height = lblAccessing.Height + GAP + GAP
    Me.Left = Screen.Width / 2 - Me.Width / 2
    Me.Top = Screen.Height / 2 - Me.Height / 2
    linTop.X1 = 0
    linTop.X2 = Me.Width
    linTop.Y1 = 0
    linTop.Y2 = 0
    'linLeft.X1 = 1
    'linLeft.X2 = 1
    'linLeft.Y1 = -100
    'linLeft.Y1 = Me.Height
    linRight.X1 = Me.Width - Screen.TwipsPerPixelX
    linRight.X2 = Me.Width - Screen.TwipsPerPixelX
    linRight.Y1 = 0
    linRight.Y2 = Me.Height
    linBottom.X1 = 0
    linBottom.X2 = Me.Width
    linBottom.Y1 = Me.Height - Screen.TwipsPerPixelY
    linBottom.Y2 = Me.Height - Screen.TwipsPerPixelY
    Call modMSAAfyLabels.Refresh(Me)
End Sub
