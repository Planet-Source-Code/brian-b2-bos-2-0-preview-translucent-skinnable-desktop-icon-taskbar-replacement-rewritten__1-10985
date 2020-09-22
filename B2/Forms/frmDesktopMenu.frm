VERSION 5.00
Begin VB.Form frmDesktopMenu 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   ClientHeight    =   3195
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4680
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   213
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   312
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmrPopup 
      Enabled         =   0   'False
      Interval        =   250
      Left            =   1740
      Top             =   1380
   End
   Begin VB.PictureBox picOver 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   495
      Left            =   420
      ScaleHeight     =   435
      ScaleWidth      =   3795
      TabIndex        =   11
      Top             =   2400
      Visible         =   0   'False
      Width           =   3855
   End
   Begin VB.PictureBox picBorder 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   495
      Index           =   9
      Left            =   1800
      ScaleHeight     =   435
      ScaleWidth      =   1155
      TabIndex        =   10
      Top             =   2460
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.PictureBox picFile 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   0
      Left            =   0
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   309
      TabIndex        =   9
      Top             =   0
      Width           =   4635
   End
   Begin VB.PictureBox picBorder 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   255
      Index           =   8
      Left            =   -780
      ScaleHeight     =   195
      ScaleWidth      =   675
      TabIndex        =   8
      Top             =   -300
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.PictureBox picBorder 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   495
      Index           =   7
      Left            =   -2700
      ScaleHeight     =   435
      ScaleWidth      =   1155
      TabIndex        =   7
      Top             =   -2460
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.PictureBox picBorder 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   495
      Index           =   6
      Left            =   -2700
      ScaleHeight     =   435
      ScaleWidth      =   1155
      TabIndex        =   6
      Top             =   -2460
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.PictureBox picBorder 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   495
      Index           =   5
      Left            =   -2700
      ScaleHeight     =   435
      ScaleWidth      =   1155
      TabIndex        =   5
      Top             =   -2460
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.PictureBox picBorder 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   495
      Index           =   4
      Left            =   -2700
      ScaleHeight     =   435
      ScaleWidth      =   1155
      TabIndex        =   4
      Top             =   -2460
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.PictureBox picBorder 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   495
      Index           =   3
      Left            =   -2700
      ScaleHeight     =   435
      ScaleWidth      =   1155
      TabIndex        =   3
      Top             =   -2460
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.PictureBox picBorder 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   495
      Index           =   2
      Left            =   -2700
      ScaleHeight     =   435
      ScaleWidth      =   1155
      TabIndex        =   2
      Top             =   -2460
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.PictureBox picBorder 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   495
      Index           =   1
      Left            =   -2700
      ScaleHeight     =   435
      ScaleWidth      =   1155
      TabIndex        =   1
      Top             =   -2460
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.PictureBox picBorder 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   495
      Index           =   0
      Left            =   -1860
      ScaleHeight     =   435
      ScaleWidth      =   1155
      TabIndex        =   0
      Top             =   -1380
      Visible         =   0   'False
      Width           =   1215
   End
End
Attribute VB_Name = "frmDesktopMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim PxH As Integer
Dim PxW As Integer
Dim Contents As Variant
Dim TxtWidth As Integer
Dim MaxWidth As Integer
Dim CurrentPath As String
Dim CurrentIndex As Integer
Dim OldIndex As Integer
'Dim PopupIndex As Integer
'Private CurrentPopup As Integer
'Dim CurrentFormIndex As Integer

Public Sub DrawMenu(path As String)
    CurrentIndex = -1
    OldIndex = -1
    
    PopupIndex = -1
    CurrentPopup = -1
    
    CurrentPath = AddASlash(path)
    
    Contents = ListFolderItems(CurrentPath)
    Me.Height = ((UBound(Contents) + 1) * 16 + 3) * Screen.TwipsPerPixelY
    Me.top = Screen.Height - frmTaskbar.Height - Me.Height
    Me.left = frmTaskbar.picDesktopButton.left * Screen.TwipsPerPixelX
    
    For i = 0 To UBound(Contents)
        TxtWidth = picFile(0).TextWidth(RemoveExtention(Contents(i)))
        If TxtWidth > MaxWidth Then MaxWidth = TxtWidth
    Next
    Me.Width = (MaxWidth + 30) * Screen.TwipsPerPixelX
    
    LoadSkinImages
    DrawBackground
    
    If picFile.UBound > UBound(Contents) + 1 Then
        For i = UBound(Contents) To picFile.UBound - 1
            picFile(i).Visible = False
        Next
    Else
        For i = picFile.UBound + 1 To UBound(Contents)
                Load picFile(i)
        Next
    End If
    
    DisplayItem (0)
    For i = 1 To UBound(Contents)
        DisplayItem (i)
    Next
    Me.Show
    Me.Refresh
End Sub


Sub LoadSkinImages()
    picBorder(0).Picture = LoadPicture(GetSkinImage("Desktop Icons\TopBorder.bmp"))
    picBorder(1).Picture = LoadPicture(GetSkinImage("Desktop Icons\RightBorder.bmp"))
    picBorder(2).Picture = LoadPicture(GetSkinImage("Desktop Icons\BottomBorder.bmp"))
    picBorder(3).Picture = LoadPicture(GetSkinImage("Desktop Icons\LeftBorder.bmp"))
    picBorder(4).Picture = LoadPicture(GetSkinImage("Desktop Icons\TopLeft.bmp"))
    picBorder(5).Picture = LoadPicture(GetSkinImage("Desktop Icons\TopRight.bmp"))
    picBorder(6).Picture = LoadPicture(GetSkinImage("Desktop Icons\BottomRight.bmp"))
    picBorder(7).Picture = LoadPicture(GetSkinImage("Desktop Icons\BottomLeft.bmp"))
    picBorder(8).Picture = LoadPicture(GetSkinImage("Desktop Icons\MiddleStretch.bmp"))
    picBorder(9).Picture = LoadPicture(GetSkinImage("Desktop Icons\MiddleOver.bmp"))
End Sub

Sub DrawBackground()
    PxH = Me.Height / Screen.TwipsPerPixelY
    PxW = Me.Width / Screen.TwipsPerPixelX
    
    Me.Cls
    
    'Draw the border
    StretchPic 0, 1, 0, 198, 1, PxW - 2, 1
    StretchPic 1, PxW - 1, 1, 1, 198, 1, PxH - 2
    StretchPic 2, 1, PxH - 1, 198, 1, PxW - 2, 1
    StretchPic 3, 0, 1, 1, 198, 1, PxH - 2
    
    'Draw the corners
    CopyPic 4, 0, 0, 1, 1
    CopyPic 5, PxW - 1, 0, 1, 1
    CopyPic 6, PxW - 1, PxH - 1, 1, 1
    CopyPic 7, 0, PxH - 1, 1, 1
        
    'Draw the background
    StretchPic 8, 1, 1, 198, 1, PxW - 2, PxH - 2
    
    'Draw the "over" image
    picOver.Width = PxW + 10
    StretchBlt picOver.hdc, 0, 0, PxW, 17, picBorder(9).hdc, 0, 0, 198, 1, vbSrcCopy
End Sub

Sub CopyPic(Index As Integer, X As Integer, Y As Integer, Width As Integer, Height As Integer)
    BitBlt Me.hdc, X, Y, Width, Height, picBorder(Index).hdc, 0, 0, vbSrcCopy
End Sub

Sub StretchPic(Index As Integer, X As Integer, Y As Integer, Width As Integer, Height As Integer, NewWidth As Integer, NewHeight As Integer)
    StretchBlt Me.hdc, X, Y, NewWidth, NewHeight, picBorder(Index).hdc, 0, 0, Width, Height, vbSrcCopy
End Sub

Function ListFolderItems(ByVal path As String) As Variant
    'returns an array of directory names
    On Error Resume Next
    Dim Count, Items(), i, ItemName ' Declare variables.
    ItemName = Dir(path, vbDirectory Or vbArchive Or vbSystem Or vbReadOnly) ' Get first directory name.
    Count = 0

    Do While Not ItemName = ""
        'A file or directory name was returned
        If Not ItemName = "." And Not ItemName = ".." Then
            ReDim Preserve Items(Count + 1)
            Items(Count) = ItemName ' Add directory name to array
            Count = Count + 1
        End If
        ItemName = Dir ' Get another item name
    Loop
    ReDim Preserve Items(Count - 1)
    ListFolderItems = Items
End Function

Sub DisplayItem(Index As Integer, Optional Over As Boolean = False)
    picFile(Index).top = 16 * Index + 1
    picFile(Index).Width = MaxWidth + 28
    picFile(Index).Visible = True
    If Over Then
        BitBlt picFile(Index).hdc, 0, 0, PxW, 16, picOver.hdc, 0, 0, vbSrcCopy
    Else
        BitBlt picFile(Index).hdc, 0, 0, PxW, 16, Me.hdc, 1, 1, vbSrcCopy
    End If
    picFile(Index).CurrentX = 20
    picFile(Index).CurrentY = 0
    picFile(Index).Print RemoveExtention(Contents(Index))
    DrawFileIcon CurrentPath & Contents(Index), picFile(Index).hdc, Icon_Small
End Sub


Private Sub picFile_Click(Index As Integer)
    ShellFile (CurrentPath & Contents(Index))
    ToggleDesktopIcons
End Sub

Private Sub picFile_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If OldIndex <> -1 And OldIndex <> Index Then
    Out (OldIndex)
End If

If Index <> -1 And CurrentIndex <> Index Then
    Over (Index)
'    If IsDir(CurrentPath & Contents(index)) Then
'        PopupIndex = index
'        tmrPopup.Enabled = True
'    Else
'        PopupIndex = -1
'        tmrPopup_Timer
'    End If
    OldIndex = Index
    CurrentIndex = Index
End If
End Sub

Sub Over(Index As Integer)
    picFile(Index).Cls
    DisplayItem Index, True
End Sub

Sub Out(Index As Integer)
    picFile(Index).Cls
    DisplayItem Index, False
End Sub

Private Function AddASlash(path As String)
    If Right(path, 1) = "\" Then AddASlash = path Else AddASlash = path & "\"
End Function

Private Function RemoveExtention(ByVal file As String) As String
    file = StrReverse(file)
    pos = InStr(file, ".")
    file = Right(file, Len(file) - pos)
    file = StrReverse(file)
    RemoveExtention = file
End Function

Public Function HideMenu()
    'If CurrentPopup <> -1 Then Forms(CurrentFormIndex + 1).HideMenu
    If CurrentIndex <> -1 Then Out (CurrentIndex)
    If OldIndex <> -1 Then Out (OldIndex)
    Unload Me
End Function

Private Function IsDir(ByVal file As String) As Boolean
    IsDir = (Dir(AddASlash(file)) <> "")
End Function

'Private Sub tmrPopup_Timer()
'    Debug.Print "Event " & Now
'    If CurrentPopup <> -1 Then
'        Forms(CurrentFormIndex + 1).HideMenu
'        CurrentPopup = -1
'    End If
'    If PopupIndex <> -1 Then
'        Dim f As New frmDesktopMenu
'        f.DrawMenu CurrentPath & Contents(PopupIndex)
'        Me.Show
'        f.Top = Me.Top + ((PopupIndex * 16 + 1) * Screen.TwipsPerPixelY)
'        f.Left = Me.Left + Me.Width - 20
'        CurrentPopup = PopupIndex
'    End If
'    tmrPopup.Enabled = False
'
'End Sub

