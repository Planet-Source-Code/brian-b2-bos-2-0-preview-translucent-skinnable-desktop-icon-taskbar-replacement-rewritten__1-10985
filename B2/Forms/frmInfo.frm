VERSION 5.00
Begin VB.Form frmInfo 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   1530
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   5460
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
   ScaleHeight     =   102
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   364
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picTitleBar 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   315
      Left            =   0
      ScaleHeight     =   315
      ScaleWidth      =   5535
      TabIndex        =   2
      Top             =   0
      Width           =   5535
      Begin VB.Label lblTitle 
         BackColor       =   &H00000000&
         Caption         =   "[Box Title]"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   60
         TabIndex        =   3
         Top             =   60
         Width           =   5355
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   4380
      TabIndex        =   1
      Top             =   540
      Width           =   915
   End
   Begin VB.Image imgIcon 
      Height          =   480
      Left            =   180
      Picture         =   "frmInfo.frx":0000
      Top             =   660
      Width           =   480
   End
   Begin VB.Label lblInfo 
      BackColor       =   &H00FFFFFF&
      Caption         =   "[Box Label]"
      Height          =   975
      Left            =   780
      TabIndex        =   0
      Top             =   420
      Width           =   3375
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim DragNow As Boolean, DragX As Integer, DragY As Integer

Public Sub DisplayInfo(Text As String, Optional caption As String = "Information")
    lblInfo.caption = Text
    lblTitle.caption = caption
    Me.Show
End Sub

Private Sub cmdOK_Click()
    Me.Hide
    HideToolTip
End Sub

Private Sub cmdOK_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If TipOver = False Then
        ToolTipEX "Click here to close the information box.", (Me.Top / Screen.TwipsPerPixelY) + cmdOK.Top - 23, (Me.Left / Screen.TwipsPerPixelX) + cmdOK.Left + 2
    End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    HideToolTip
End Sub

Private Sub lblTitle_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    DragNow = True
    DragX = X
    DragY = Y
End Sub

Private Sub lblTitle_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If DragNow Then
        Me.Top = Me.Top + Y - DragY
        Me.Left = Me.Left + X - DragX
    End If
End Sub

Private Sub lblTitle_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    DragNow = False
End Sub
