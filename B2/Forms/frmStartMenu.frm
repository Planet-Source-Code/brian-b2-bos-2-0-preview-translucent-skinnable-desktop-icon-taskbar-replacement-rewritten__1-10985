VERSION 5.00
Begin VB.Form frmStartMenu 
   BorderStyle     =   0  'None
   ClientHeight    =   4800
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6900
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   320
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   460
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picDesktopCapture 
      AutoRedraw      =   -1  'True
      Height          =   4875
      Left            =   0
      ScaleHeight     =   4815
      ScaleWidth      =   3915
      TabIndex        =   16
      Top             =   0
      Visible         =   0   'False
      Width           =   3975
   End
   Begin VB.PictureBox picItemOver 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C0C0FF&
      BorderStyle     =   0  'None
      Height          =   600
      Index           =   7
      Left            =   3300
      ScaleHeight     =   600
      ScaleWidth      =   3315
      TabIndex        =   15
      Top             =   0
      Width           =   3315
   End
   Begin VB.PictureBox picItemOver 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   0  'None
      Height          =   600
      Index           =   6
      Left            =   3300
      ScaleHeight     =   600
      ScaleWidth      =   3315
      TabIndex        =   14
      Top             =   600
      Width           =   3315
   End
   Begin VB.PictureBox picItemOver 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C0C0FF&
      BorderStyle     =   0  'None
      Height          =   600
      Index           =   5
      Left            =   3300
      ScaleHeight     =   600
      ScaleWidth      =   3315
      TabIndex        =   13
      Top             =   1200
      Width           =   3315
   End
   Begin VB.PictureBox picItemOver 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   0  'None
      Height          =   600
      Index           =   4
      Left            =   3300
      ScaleHeight     =   600
      ScaleWidth      =   3315
      TabIndex        =   12
      Top             =   1800
      Width           =   3315
   End
   Begin VB.PictureBox picItemOver 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C0C0FF&
      BorderStyle     =   0  'None
      Height          =   600
      Index           =   3
      Left            =   3300
      ScaleHeight     =   600
      ScaleWidth      =   3315
      TabIndex        =   11
      Top             =   2400
      Width           =   3315
   End
   Begin VB.PictureBox picItemOver 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   0  'None
      Height          =   600
      Index           =   2
      Left            =   3300
      ScaleHeight     =   600
      ScaleWidth      =   3315
      TabIndex        =   10
      Top             =   3000
      Width           =   3315
   End
   Begin VB.PictureBox picItemOver 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   0  'None
      Height          =   600
      Index           =   0
      Left            =   3300
      ScaleHeight     =   600
      ScaleWidth      =   3315
      TabIndex        =   9
      Top             =   4200
      Width           =   3315
   End
   Begin VB.PictureBox picItemOver 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C0C0FF&
      BorderStyle     =   0  'None
      Height          =   600
      Index           =   1
      Left            =   3300
      ScaleHeight     =   600
      ScaleWidth      =   3315
      TabIndex        =   8
      Top             =   3600
      Width           =   3315
   End
   Begin VB.PictureBox picItem 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   0  'None
      Height          =   600
      Index           =   7
      Left            =   0
      ScaleHeight     =   600
      ScaleWidth      =   3315
      TabIndex        =   7
      Top             =   0
      Width           =   3315
   End
   Begin VB.PictureBox picItem 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Height          =   600
      Index           =   6
      Left            =   0
      ScaleHeight     =   600
      ScaleWidth      =   3315
      TabIndex        =   6
      Top             =   600
      Width           =   3315
   End
   Begin VB.PictureBox picItem 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   0  'None
      Height          =   600
      Index           =   5
      Left            =   0
      ScaleHeight     =   600
      ScaleWidth      =   3315
      TabIndex        =   5
      Top             =   1200
      Width           =   3315
   End
   Begin VB.PictureBox picItem 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Height          =   600
      Index           =   4
      Left            =   0
      ScaleHeight     =   600
      ScaleWidth      =   3315
      TabIndex        =   4
      Top             =   1800
      Width           =   3315
   End
   Begin VB.PictureBox picItem 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   0  'None
      Height          =   600
      Index           =   3
      Left            =   0
      ScaleHeight     =   600
      ScaleWidth      =   3315
      TabIndex        =   3
      Top             =   2400
      Width           =   3315
   End
   Begin VB.PictureBox picItem 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Height          =   600
      Index           =   2
      Left            =   0
      ScaleHeight     =   600
      ScaleWidth      =   3315
      TabIndex        =   2
      Top             =   3000
      Width           =   3315
   End
   Begin VB.PictureBox picItem 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   0  'None
      Height          =   600
      Index           =   1
      Left            =   0
      ScaleHeight     =   600
      ScaleWidth      =   3315
      TabIndex        =   1
      Top             =   3600
      Width           =   3315
   End
   Begin VB.PictureBox picItem 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Height          =   600
      Index           =   0
      Left            =   0
      ScaleHeight     =   600
      ScaleWidth      =   3315
      TabIndex        =   0
      Top             =   4200
      Width           =   3315
   End
End
Attribute VB_Name = "frmStartMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public StartMenuHeight As Integer
Public StartMenuWidth As Integer

Dim CurrentIndex As Integer, OldIndex As Integer

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyUp
            If CurrentIndex = 7 Then CurrentIndex = 0 Else CurrentIndex = CurrentIndex + 1
                SelectImage (CurrentIndex)
                If OldIndex <> -1 Then UnselectImage (OldIndex)
                OldIndex = CurrentIndex
        Case vbKeyDown
            If CurrentIndex = 0 Then CurrentIndex = 7 Else CurrentIndex = CurrentIndex - 1
                SelectImage (CurrentIndex)
                If OldIndex <> -1 Then UnselectImage (OldIndex)
                OldIndex = CurrentIndex
        Case vbKeyEscape
            HideMe
    End Select
End Sub

Private Sub Form_Load()
    LoadSkinSettings
    LoadSkinImages
    
    OldIndex = -1
    CurrentIndex = -1
End Sub

Sub Display()
    BltDesktop 0, ScreenHeight - StartMenuHeight - frmTaskbar.TaskbarHeight, picDesktopCapture
    For i = 0 To 7
        picItem(i).Cls
        AlphaBlending picItem(i).HDC, 0, 0, StartMenuWidth, StartMenuHeight / 8, picDesktopCapture.HDC, 0, (7 - i) * (StartMenuHeight / 8), StartMenuWidth, StartMenuHeight / 8, 50
    Next
    
    SetWindowPos Me.hWND, HWND_TOPMOST, 0, ScreenHeight - StartMenuHeight - frmTaskbar.TaskbarHeight, StartMenuWidth, StartMenuHeight, 0
    Me.Show
    Me.Refresh
End Sub

Sub HideMe()
    If OldIndex <> -1 Then UnselectImage (OldIndex)
    If CurrentIndex <> -1 Then UnselectImage (CurrentIndex)
    CurrentIndex = -1
    OldIndex = -1
    Me.Hide
End Sub

Sub LoadSkinSettings()
    StartMenuWidth = GetStartMenuWidth
    StartMenuHeight = GetStartMenuHeight
End Sub

Sub LoadSkinImages()
    For i = 0 To 7
        picItem(i).Picture = LoadPicture(GetSkinImage("Start Menu\Main Menu\Normal\Image" & i & ".bmp"))
        picItemOver(i).Picture = LoadPicture(GetSkinImage("Start Menu\Main Menu\Over\Image" & i & ".bmp"))
    Next
End Sub

Sub UnselectImage(Index As Integer)
    picItem(Index).Cls
    AlphaBlending picItem(Index).HDC, 0, 0, StartMenuWidth, StartMenuHeight / 8, picDesktopCapture.HDC, 0, (7 - Index) * (StartMenuHeight / 8), StartMenuWidth, StartMenuHeight / 8, 50
    picItem(Index).Refresh
    OldIndex = -1
End Sub

Sub SelectImage(Index As Integer)
    BitBlt picItem(Index).HDC, 0, 0, StartMenuWidth, StartMenuHeight / 8, picItemOver(Index).HDC, 0, 0, vbSrcCopy
    AlphaBlending picItem(Index).HDC, 0, 0, StartMenuWidth, StartMenuHeight / 8, picDesktopCapture.HDC, 0, (7 - Index) * (StartMenuHeight / 8), StartMenuWidth, StartMenuHeight / 8, 50
    picItem(Index).Refresh
End Sub

Private Sub picItem_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If OldIndex <> -1 And OldIndex <> Index Then UnselectImage (OldIndex)
    If CurrentIndex <> Index Then SelectImage (Index)
    OldIndex = Index
    CurrentIndex = Index
End Sub
