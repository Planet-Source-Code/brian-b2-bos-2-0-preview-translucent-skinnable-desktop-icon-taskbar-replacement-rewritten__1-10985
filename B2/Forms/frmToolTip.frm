VERSION 5.00
Begin VB.Form frmToolTip 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   360
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   6855
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   24
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   457
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picDesktopCapture 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   435
      Left            =   2820
      ScaleHeight     =   435
      ScaleWidth      =   1215
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label lblTip 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "[ Tool Tip Text ]"
      Height          =   255
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   4395
   End
End
Attribute VB_Name = "frmToolTip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private OldHwnd As Long

Public Function DisplayTip(Text As String, Top As Integer, Left As Integer)
    lblTip.caption = Text
    If Me.TextWidth(Text) + 8 > lblTip.Width Then
        Width = lblTip.Width + 8
        Height = (24 * Int(Me.TextWidth(Text) / lblTip.Width)) + 12
        lblTip.Height = Height
        Top = Top - Me.Height + 24
    Else
        Height = 24
        Width = Me.TextWidth(Text) + 8
    End If
    SetWindowPos Me.hwnd, -1, Left, Top, Width, Height, SWP_SHOWWINDOW Or SWP_NOACTIVATE
    picDesktopCapture.Width = Width
    picDesktopCapture.Height = Height
    Me.Cls
    BltDesktop Left + 1, Top + 1, picDesktopCapture
    AlphaBlending Me.HDC, 0, 0, Width, Height, picDesktopCapture.HDC, 0, 0, Width, Height, 80
    'Move the window into place and bring it to the top
    DisplayTip = True
    ToolTipDisplayed = True
End Function

Public Sub HideTip()
    Me.Hide
    ToolTipDisplayed = False
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    HideTip
End Sub

Private Sub lblTip_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    HideTip
End Sub

