VERSION 5.00
Begin VB.Form frmTaskbar 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   5280
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9915
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   352
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   661
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picDesktopCapture 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   4500
      ScaleHeight     =   495
      ScaleWidth      =   1215
      TabIndex        =   16
      Top             =   0
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.PictureBox picQuickstart 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   480
      Left            =   900
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   84
      TabIndex        =   15
      Top             =   0
      Width           =   1260
      Begin VB.Image imgIcon 
         Height          =   255
         Index           =   0
         Left            =   45
         Top             =   120
         Width           =   255
      End
   End
   Begin VB.Timer tmrCheckFocus 
      Interval        =   1
      Left            =   4920
      Top             =   2400
   End
   Begin VB.Timer tmrUpdateTime 
      Interval        =   5000
      Left            =   4380
      Top             =   2400
   End
   Begin VB.PictureBox picTaskbarImage 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   495
      Index           =   8
      Left            =   3000
      ScaleHeight     =   495
      ScaleWidth      =   1215
      TabIndex        =   13
      Top             =   2520
      Width           =   1215
   End
   Begin VB.PictureBox picTaskbarImage 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   495
      Index           =   7
      Left            =   420
      ScaleHeight     =   495
      ScaleWidth      =   3375
      TabIndex        =   12
      Top             =   2640
      Visible         =   0   'False
      Width           =   3375
   End
   Begin VB.PictureBox picTaskbarImage 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   495
      Index           =   6
      Left            =   780
      ScaleHeight     =   495
      ScaleWidth      =   3375
      TabIndex        =   11
      Top             =   3660
      Visible         =   0   'False
      Width           =   3375
   End
   Begin VB.PictureBox picTray 
      Align           =   4  'Align Right
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   0  'None
      Height          =   5280
      Left            =   8415
      ScaleHeight     =   352
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   100
      TabIndex        =   10
      Top             =   0
      Width           =   1500
      Begin VB.Label lblTime 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "3:00 pm"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   660
         TabIndex        =   14
         Top             =   120
         Width           =   720
      End
   End
   Begin VB.PictureBox picTaskbarImage 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   495
      Index           =   5
      Left            =   300
      ScaleHeight     =   495
      ScaleWidth      =   3375
      TabIndex        =   8
      Top             =   3360
      Visible         =   0   'False
      Width           =   3375
   End
   Begin VB.PictureBox picProgram 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   450
      Index           =   0
      Left            =   2100
      OLEDropMode     =   1  'Manual
      ScaleHeight     =   30
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   137
      TabIndex        =   7
      Top             =   0
      Width           =   2055
      Begin VB.Label lblCaption 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   300
         OLEDropMode     =   1  'Manual
         TabIndex        =   9
         Top             =   120
         Width           =   1215
      End
   End
   Begin VB.PictureBox picTaskbarImage 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   495
      Index           =   4
      Left            =   60
      ScaleHeight     =   495
      ScaleWidth      =   1215
      TabIndex        =   6
      Top             =   3120
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.PictureBox picDesktopButton 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   420
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   30
      TabIndex        =   5
      Top             =   0
      Width           =   450
   End
   Begin VB.PictureBox picStartButton 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   0
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   30
      TabIndex        =   4
      Top             =   0
      Width           =   450
   End
   Begin VB.PictureBox picTaskbarImage 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   495
      Index           =   3
      Left            =   240
      ScaleHeight     =   495
      ScaleWidth      =   1215
      TabIndex        =   3
      Top             =   2580
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.PictureBox picTaskbarImage 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   495
      Index           =   2
      Left            =   60
      ScaleHeight     =   495
      ScaleWidth      =   1215
      TabIndex        =   2
      Top             =   1740
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.PictureBox picTaskbarImage 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   495
      Index           =   1
      Left            =   120
      ScaleHeight     =   495
      ScaleWidth      =   1215
      TabIndex        =   1
      Top             =   1380
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.PictureBox picTaskbarImage 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   495
      Index           =   0
      Left            =   120
      ScaleHeight     =   495
      ScaleWidth      =   1215
      TabIndex        =   0
      Top             =   780
      Width           =   1215
   End
End
Attribute VB_Name = "frmTaskbar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const TranslucencyLevel = 60

'Declerations for the "App-Tray"
Private Declare Function OleCreatePictureIndirect Lib "oleaut32.dll" _
        (pDicDesc As IconType, riid As CLSIdType, ByVal fown As Long, _
        lpUnk As Object) As Long
        
Private Declare Function SHGetFileInfo Lib "shell32.dll" Alias _
        "SHGetFileInfoA" (ByVal pszPath As String, ByVal _
        dwFileAttributes As Long, psfi As ShellFileInfoType, ByVal _
        cbFileInfo As Long, ByVal uFlags As Long) As Long

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWND As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Type IconType
    cbSize As Long
    picType As PictureTypeConstants
    hIcon As Long
End Type

Private Type CLSIdType
    id(16) As Byte
End Type

Private Type ShellFileInfoType
  hIcon As Long
  iIcon As Long
  dwAttributes As Long
  szDisplayName As String * 260
  szTypeName As String * 80
End Type

Const Large = &H100
Const Small = &H101

'Taskbar Settings (loaded from skin)
Public TaskbarHeight As Integer 'Height of the taskbar (in pixels)
Public StartButtonWidth As Integer 'Width of the start button (in pixels)
Public DesktopButtonWidth As Integer 'Width of the "desktop popup" button
Public TaskbarButtonWidth As Integer

'Taskbar Settings (determined automaticlly)
Public TaskbarWidth As Integer 'Width of the taskbar (in pixels)
Public TaskbarButtons As Integer
Public MaxTaskLength As Integer

'Program list variables
Dim ButtonVisible() As Boolean
Dim ButtonDown() As Boolean
Dim ButtonCaption() As String
Dim AppTrayPath As String 'Path for the "quickstart" menu
Dim AppTrayCurrentFilename As String, Number1 As Integer, Index1 As Integer
Dim AppTrayOverIndex As Integer
Dim ProgName As String

Dim ProgramOverIndex As Integer

Dim StartButtonOver As Boolean
Dim DesktopButtonOver As Boolean

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then UpdateIcons
End Sub

Private Sub Form_Load()
    AppTrayPath = App.path & "\quickstart\"   'Set Path to the Quickstart programs
    LoadSkinSettings 'Get taskbar settings determined by the skin
    GetTaskbarSettings 'Get taskbar settings not determined by the skin
    LoadSkinImages 'Load the pictures for the taskbar
    
    picDesktopCapture.Width = ScreenWidth: picDesktopCapture.Height = ScreenHeight
    BltDesktop 0, ScreenHeight - TaskbarHeight, picDesktopCapture
    
    ApplyTaskbarSettings 'Apply the taskbar settings
End Sub


Sub GetTaskbarSettings()
    TaskbarWidth = ScreenWidth
    TaskbarButtons = Int((TaskbarWidth - DesktopButtonWidth - StartButtonWidth) / TaskbarButtonWidth) - 1
    MatchFonts
    For i = 10 To 100
        If picTaskbarImage(0).TextWidth(Space(i)) > TaskbarWidth - 20 Then Exit For
    Next
    MaxTaskLength = i
    ReDim ButtonVisible(TaskbarButtons)
    ReDim ButtonDown(TaskbarButtons)
    ReDim ButtonCaption(TaskbarButtons)
End Sub

Sub MatchFonts()
    picTaskbarImage(0).FontName = lblCaption(0).FontName
    picTaskbarImage(0).FontSize = lblCaption(0).FontSize
    picTaskbarImage(0).FontBold = lblCaption(0).FontBold
End Sub

Sub LoadSkinSettings()
    TaskbarHeight = GetTaskbarHeight
    StartButtonWidth = GetStartButtonWidth
    DesktopButtonWidth = GetDesktopButtonWidth
    TaskbarButtonWidth = GetTaskbarButtonWidth
End Sub

Sub LoadSkinImages()
    picProgram(0).Picture = LoadPicture(GetSkinImage("Taskbar\Program.bmp")) 'Load the picture for the program button
    picTaskbarImage(0).Picture = LoadPicture(GetSkinImage("Taskbar\TaskbarBG.bmp")) 'Load the picture for the taskbar background
    picTaskbarImage(1).Picture = LoadPicture(GetSkinImage("Taskbar\startbutton.bmp")) 'Load the picture for the start button
    picTaskbarImage(2).Picture = LoadPicture(GetSkinImage("Taskbar\startbuttondown.bmp")) 'Load the picture for the pressed start button
    picTaskbarImage(3).Picture = LoadPicture(GetSkinImage("Taskbar\desktop.bmp")) 'Load the icon for the desktop popup
    picTaskbarImage(4).Picture = LoadPicture(GetSkinImage("Taskbar\desktopdown.bmp")) 'Load the icon for the pressed dekstop popup
    picTaskbarImage(5).Picture = LoadPicture(GetSkinImage("Taskbar\ProgramDown.bmp")) 'Load the picture for the pressed program button
    picTaskbarImage(6).Picture = LoadPicture(GetSkinImage("Taskbar\SystemTray\Left.bmp")) 'Load the picture for the left side of the tray
    picTaskbarImage(7).Picture = LoadPicture(GetSkinImage("Taskbar\SystemTray\Center.bmp")) 'Load the picture to be streatched across the middle of the tray
    picTaskbarImage(8).Picture = LoadPicture(GetSkinImage("Taskbar\SystemTray\Right.bmp")) 'Load the picture for the right side of the tray
End Sub

Sub ApplyTaskbarSettings()
    SetWindowPos Me.hWND, -1, 0, ScreenHeight - TaskbarHeight, ScreenWidth, TaskbarHeight, 0  'Move the window into place and bring it to the top
    
    StretchBlt Me.HDC, 0, 0, ScreenWidth, TaskbarHeight, picTaskbarImage(0).HDC, 0, 0, picTaskbarImage(0).Width, TaskbarHeight, vbSrcCopy
    
    picStartButton.Height = TaskbarHeight 'Set the height of the start button
    picStartButton.Width = StartButtonWidth 'Set the width of the start button
    BitBlt picStartButton.HDC, 0, 0, StartButtonWidth, TaskbarHeight, picTaskbarImage(1).HDC, 0, 0, vbSrcCopy 'Copy the image for the start button

    picDesktopButton.Height = TaskbarHeight 'Set the height of the desktop popup icon
    picDesktopButton.Left = StartButtonWidth
    picDesktopButton.Width = DesktopButtonWidth 'Set the width of the desktop popup icon
    BitBlt picDesktopButton.HDC, 0, 0, DesktopButtonWidth, TaskbarHeight, picTaskbarImage(3).HDC, 0, 0, vbSrcCopy 'Copy the image for the desktop popup icon

    AlphaBlending Me.HDC, 0, 0, ScreenWidth, TaskbarHeight, picDesktopCapture.HDC, 0, 0, ScreenWidth, TaskbarHeight, TranslucencyLevel
    AlphaBlending picStartButton.HDC, 0, 0, StartButtonWidth, TaskbarHeight, picDesktopCapture.HDC, 0, 0, StartButtonWidth, TaskbarHeight, TranslucencyLevel
    AlphaBlending picDesktopButton.HDC, 0, 0, DesktopButtonWidth, TaskbarHeight, picDesktopCapture.HDC, StartButtonWidth, 0, DesktopButtonWidth, TaskbarHeight, TranslucencyLevel

    UpdateIcons
    
    lblCaption(0).Width = TaskbarButtonWidth - 20
    picProgram(0).Width = TaskbarButtonWidth
    picProgram(0).Left = picQuickstart.Left + picQuickstart.Width + 2
    For i = 1 To TaskbarButtons
        Load picProgram(i)
        picProgram(i).Left = picProgram(i - 1).Left + picProgram(i - 1).Width
        picProgram(i).Visible = True
        picProgram(i).Width = TaskbarButtonWidth
        
        Load lblCaption(i)
        Set lblCaption(i).Container = picProgram(i)
        lblCaption(i).Width = TaskbarButtonWidth - 20
        lblCaption(i).Visible = True
    Next
    UpdateTime
    ResizeTray 100
End Sub

Sub UpdateIcons()
AppTrayOverIndex = -1
If AppTrayPath = "" Then
picQuickstart.Width = 0
picQuickstart.Visible = False
Exit Sub
End If
AppTrayCurrentFilename = Dir(AppTrayPath, vbNormal)   'Get first file
If AppTrayCurrentFilename <> "" Then
    If Number1 > 0 Then
    For n = 1 To Number1
        Unload imgIcon(n)
        picQuickstart.Picture = LoadPicture()
    Next n
End If

Number1 = -1
Do While AppTrayCurrentFilename <> ""
   'Ignore actual and higher directory
   If AppTrayCurrentFilename <> "." And AppTrayCurrentFilename <> ".." Then
      'Be sure that AppTrayCurrentFilename is not a directory
      If (GetAttr(AppTrayPath & AppTrayCurrentFilename)) <> vbDirectory Then
      Number1 = Number1 + 1
        If Number1 > 0 Then
        Load imgIcon(Number1)
        imgIcon(Number1).Left = imgIcon(Number1 - 1).Left + imgIcon(Number1 - 1).Width + 3
        imgIcon(Number1).Picture = LoadIcon(Small)
        imgIcon(Number1).Tag = AppTrayCurrentFilename
        imgIcon(Number1).Visible = True
        Else
        imgIcon(0).Picture = LoadIcon(Small)
        imgIcon(0).Tag = AppTrayCurrentFilename
        End If
      End If
   End If
   AppTrayCurrentFilename = Dir   'Get next file
Loop
picQuickstart.Width = imgIcon(Number1).Left + imgIcon(Number1).Width + 5

'Set background (needed for updating icons)
BitBlt picQuickstart.HDC, 0, 0, 5, TaskbarHeight, picTaskbarImage(6).HDC, 0, 0, vbSrcCopy
StretchBlt picQuickstart.HDC, 5, 0, picQuickstart.Width, TaskbarHeight, picTaskbarImage(7).HDC, 0, 0, 1, TaskbarHeight, vbSrcCopy
BitBlt picQuickstart.HDC, picQuickstart.Width - 5, 0, 5, TaskbarHeight, picTaskbarImage(8).HDC, 0, 0, vbSrcCopy
AlphaBlending picQuickstart.HDC, 0, 0, picQuickstart.Width, TaskbarHeight, picDesktopCapture.HDC, picQuickstart.Left, 0, picQuickstart.Width, TaskbarHeight, TranslucencyLevel
Else
'If there are no files present don't show the "App-Tray"
picQuickstart.Width = 0
picQuickstart.Visible = False
End If
'Set the positions for the program buttons
For n = 0 To picProgram.Count - 1
If n = 0 Then
picProgram(n).Left = picQuickstart.Left + picQuickstart.Width + 2
Else
picProgram(n).Left = picProgram(n - 1).Left + picProgram(n - 1).Width
End If
Next n
End Sub

'Get the icons for the "App-Tray
Private Function LoadIcon(Size&) As IPictureDisp
  Dim Result&, file$, Slash$
  Dim Unkown As IUnknown
  Dim Icon As IconType
  Dim CLSID As CLSIdType
  Dim ShellInfo As ShellFileInfoType
    
    file = AppTrayPath & AppTrayCurrentFilename
    Call SHGetFileInfo(file, 0, ShellInfo, Len(ShellInfo), Size)
 
    Icon.cbSize = Len(Icon)
    Icon.picType = vbPicTypeIcon
    Icon.hIcon = ShellInfo.hIcon
    CLSID.id(8) = &HC0
    CLSID.id(15) = &H46
    Result = OleCreatePictureIndirect(Icon, CLSID, 1, Unkown)
    
    Set LoadIcon = Unkown
End Function

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    HideTaskbarTips
End Sub

'Open program in the "App-Tray"
Private Sub imgIcon_DblClick(Index As Integer)

End Sub

Private Sub imgIcon_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
imgIcon(Index).Top = imgIcon(Index).Top + 1
imgIcon(Index).Left = imgIcon(Index).Left + 1
End Sub

Private Sub imgIcon_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If ProgramOverIndex <> -1 Then ProgramOverIndex = -1: HideToolTip
    If AppTrayOverIndex <> Index Then
        StartButtonOver = False
        HideToolTip
        AppTrayOverIndex = Index
        ProgName = imgIcon(Index).Tag
        If Right(ProgName, 4) = ".lnk" Then
            ToolTipEX Left(ProgName, Len(ProgName) - 4), ScreenHeight - TaskbarHeight - 20, StartButtonWidth + DesktopButtonWidth + 16 * Index
        Else
            ToolTipEX ProgName, ScreenHeight - TaskbarHeight - 20, StartButtonWidth + DesktopButtonWidth + 16 * Index
        End If
    End If
End Sub

Private Sub imgIcon_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
ShellExecute Me.hWND, "open", AppTrayPath & imgIcon(Index).Tag, "", "", 1
imgIcon(Index).Top = imgIcon(Index).Top - 1
imgIcon(Index).Left = imgIcon(Index).Left - 1
End Sub

Private Sub lblCaption_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Temp fix
    AppActivate modWindowAPICalls.WindowName(Index)
    UpdateTaskbar
End Sub

Private Sub lblCaption_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If AppTrayOverIndex <> -1 Then AppTrayOverIndex = -1: HideToolTip
    If ProgramOverIndex <> Index Then
        StartButtonOver = False
        HideToolTip
        ProgramOverIndex = Index
        ToolTipEX modWindowAPICalls.WindowName(Index), ScreenHeight - TaskbarHeight - 24, Index * TaskbarButtonWidth + StartButtonWidth + DesktopButtonWidth + picQuickstart.Width
    End If
End Sub

Private Sub lblCaption_OLEDragDrop(Index As Integer, Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    frmInfo.DisplayInfo "You cannot drag an item onto a taskbar button. However, if you drag an item over a taskbar button, the program that that button represents will come to the front.", "Taskbar drag problem"
End Sub

Private Sub lblCaption_OLEDragOver(Index As Integer, Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, State As Integer)
    lblCaption_MouseDown Index, 0, 0, 0, 0
End Sub

Private Sub lblTime_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ToolTipEX Now, ScreenHeight - TaskbarHeight - 24, ScreenWidth - 150
End Sub

Private Sub picDesktopButton_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ToggleDesktopIcons
    HideToolTip
End Sub

Private Sub picDesktopButton_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If DesktopButtonOver = False Then HideTaskbarTips: DesktopButtonOver = True
    ToolTipEX "Click here to display a list of your desktop icons", ScreenHeight - TaskbarHeight - 20, StartButtonWidth
End Sub

Private Sub picProgram_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblCaption_MouseDown Index, 0, 0, 0, 0
End Sub

Private Sub picProgram_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblCaption_MouseMove Index, 0, 0, 0, 0
End Sub

Private Sub picProgram_OLEDragDrop(Index As Integer, Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    frmInfo.DisplayInfo "You cannot drag an item onto a taskbar button. However, if you drag an item over a taskbar button, the program that that button represents will come to the front.", "Taskbar drag problem"
End Sub

Private Sub picProgram_OLEDragOver(Index As Integer, Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, State As Integer)
    lblCaption_MouseDown Index, 0, 0, 0, 0
End Sub

Private Sub picStartButton_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    SetWindowPos frmToolTip.hWND, 0, 0, 0, 0, 0, SWP_HIDEWINDOW
    ToggleStartMenu
End Sub

Sub StartButtonOn()
    BitBlt picStartButton.HDC, 0, 0, StartButtonWidth, TaskbarHeight, picTaskbarImage(2).HDC, 0, 0, vbSrcCopy 'Copy the image for the start button
    AlphaBlending picStartButton.HDC, 0, 0, StartButtonWidth, TaskbarHeight, picDesktopCapture.HDC, 0, 0, StartButtonWidth, TaskbarHeight, TranslucencyLevel
    picStartButton.Refresh
End Sub

Sub StartButtonOff()
    BitBlt picStartButton.HDC, 0, 0, StartButtonWidth, TaskbarHeight, picTaskbarImage(1).HDC, 0, 0, vbSrcCopy 'Copy the image for the start button
    AlphaBlending picStartButton.HDC, 0, 0, StartButtonWidth, TaskbarHeight, picDesktopCapture.HDC, 0, 0, StartButtonWidth, TaskbarHeight, TranslucencyLevel
    picStartButton.Refresh
End Sub

Sub DesktopIconOn()
    BitBlt picDesktopButton.HDC, 0, 0, StartButtonWidth, TaskbarHeight, picTaskbarImage(4).HDC, 0, 0, vbSrcCopy 'Copy the image for the start button
    AlphaBlending picDesktopButton.HDC, 0, 0, DesktopButtonWidth, TaskbarHeight, picDesktopCapture.HDC, StartButtonWidth, 0, DesktopButtonWidth, TaskbarHeight, TranslucencyLevel
    picDesktopButton.Refresh
End Sub

Sub DesktopIconOff()
    BitBlt picDesktopButton.HDC, 0, 0, StartButtonWidth, TaskbarHeight, picTaskbarImage(3).HDC, 0, 0, vbSrcCopy 'Copy the image for the start button
    AlphaBlending picDesktopButton.HDC, 0, 0, DesktopButtonWidth, TaskbarHeight, picDesktopCapture.HDC, StartButtonWidth, 0, DesktopButtonWidth, TaskbarHeight, TranslucencyLevel
    picDesktopButton.Refresh
End Sub

Private Sub picStartButton_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If StartButtonOver = False Then HideTaskbarTips: StartButtonOver = True
    If StartMenuShown = False Then
        ToolTipEX "Click here to display the B2 Start Menu", ScreenHeight - TaskbarHeight - 20, 0
    End If
End Sub

Private Sub picTray_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
HideTaskbarTips
End Sub

Private Sub tmrCheckFocus_Timer()
    Dim WindowName As Variant, WindowHwnd As Variant, AWI As Integer
    If StartMenuShown And GetForegroundWindow <> frmStartMenu.hWND Then ToggleStartMenu
    If DesktopIconsShown And GetForegroundWindow <> frmDesktopMenu.hWND Then ToggleDesktopIcons
    'If ToolTipDisplayed And GetForegroundWindow <> frmToolTip.hWND Then HideTaskbarTips
    UpdateTaskbar
End Sub

Sub UpdateTaskbar()
    Call modWindowAPICalls.GetWindows
    For i = 0 To TaskbarButtons
        If i >= UBound(modWindowAPICalls.WindowID) Then
            If ButtonVisible(Index) Then DisableButton (i)
        Else
            If ButtonVisible(i) = False Then EnableButton (i)
            If ButtonCaption(i) <> modWindowAPICalls.WindowName(i) Then
                UpdateButtonCaption (i)
                picProgram(i).Cls
                If ButtonDown(i) Then MakeButtonDown (i) Else UpdateButtonIcon (i): MakeTaskbarButtonTranslucent (i)
            End If
            If i = modWindowAPICalls.AWI Then
                If ButtonDown(i) = False Then MakeButtonDown (i)
            Else
                If ButtonDown(i) Then MakeButtonUp (i)
            End If
        End If
    Next
End Sub

Sub DisableButton(Index As Integer)
    picProgram(Index).Visible = False
    ButtonVisible(Index) = False
End Sub

Sub EnableButton(Index As Integer)
    picProgram(Index).Visible = True
    ButtonVisible(Index) = True
    picProgram(Index).Cls
    MakeTaskbarButtonTranslucent (Index)
End Sub

Sub MakeButtonDown(Index As Integer)
    ButtonDown(Index) = True
    BitBlt picProgram(Index).HDC, 0, 0, TaskbarButtonWidth, TaskbarHeight, picTaskbarImage(5).HDC, 0, 0, vbSrcCopy
    UpdateButtonIcon (Index)
    picProgram(Index).Refresh
    lblCaption(Index).Top = 9
    lblCaption(Index).Left = 21
    MakeTaskbarButtonTranslucent (Index)
End Sub

Sub MakeButtonUp(Index As Integer)
    ButtonDown(Index) = False
    picProgram(Index).Cls
    UpdateButtonIcon (Index)
    lblCaption(Index).Top = 8
    lblCaption(Index).Left = 20
    UpdateButtonIcon (Index)
    MakeTaskbarButtonTranslucent (Index)
End Sub

Sub MakeTaskbarButtonTranslucent(Index As Integer)
    AlphaBlending picProgram(Index).HDC, 0, 0, TaskbarButtonWidth, TaskbarHeight, picDesktopCapture.HDC, picProgram(Index).Left, 0, TaskbarButtonWidth, TaskbarHeight, TranslucencyLevel
    picProgram(Index).Refresh
End Sub

Sub UpdateButtonCaption(Index As Integer)
    cap = modWindowAPICalls.WindowName(Index)
    ButtonCaption(Index) = cap
    If Len(cap) > MaxTaskLength Then
        cap = Left(cap, MaxTaskLength - 3) & "..."
    End If
    lblCaption(Index).caption = cap
End Sub

Sub UpdateButtonIcon(Index As Integer)
Down = ButtonDown(Index)
If Down Then
    DrawIcon picProgram(Index).HDC, modWindowAPICalls.WindowID(Index), 6, 7
Else
    DrawIcon picProgram(Index).HDC, modWindowAPICalls.WindowID(Index), 5, 6
End If
End Sub

Sub DrawIcon(HDC As Long, hWND As Long, X As Integer, Y As Integer, Optional largesize As Boolean = False)
        ico = GetIcon(hWND)
        If largesize Then
        
        Else
            DrawIconEx HDC, X, Y, ico, 16, 16, 0, 0, DI_NORMAL
        End If
End Sub

Sub ResizeTray(NewWidth As Integer)
    picTray.Width = NewWidth + 3
    BitBlt picTray.HDC, 0, 0, 5, TaskbarHeight, picTaskbarImage(6).HDC, 0, 0, vbSrcCopy
    StretchBlt picTray.HDC, 5, 0, NewWidth - 10, TaskbarHeight, picTaskbarImage(7).HDC, 0, 0, 1, TaskbarHeight, vbSrcCopy
    BitBlt picTray.HDC, NewWidth - 5, 0, 5, TaskbarHeight, picTaskbarImage(8).HDC, 0, 0, vbSrcCopy
    StretchBlt picTray.HDC, NewWidth, 0, 2, TaskbarHeight, picTaskbarImage(0).HDC, 0, 0, 1, TaskbarHeight, vbSrcCopy
    lblTime.Left = NewWidth - lblTime.Width - 6
    AlphaBlending picTray.HDC, 0, 0, NewWidth, TaskbarHeight, picDesktopCapture.HDC, ScreenWidth - NewWidth, 0, NewWidth, TaskbarHeight, TranslucencyLevel
End Sub

Sub UpdateTime()
    lblTime.caption = NiceTime(True)
End Sub

Function NiceTime(ampm As Boolean)
If ampm Then
    a = Hour(Now)
    If a > 12 Then a = a - 12: strampm = "pm" Else strampm = "am"
    NiceTime = a & ":" & Format(Minute(Now), "00") & " " & strampm
Else
    NiceTime = Format(Hour(Now), "00") & ":" & Format(Minute(Now), "00")
End If
End Function

Private Sub tmrUpdateTime_Timer()
    UpdateTime
    TrayWidth = 100
    lblTime.Left = TrayWidth - lblTime.Width - 6
End Sub

Sub HideTaskbarTips()
    If AppTrayOverIndex <> -1 Then HideToolTip: AppTrayOverIndex = -1
    If ProgramOverIndex <> -1 Then HideToolTip: ProgramOverIndex = -1
    If StartButtonOver = True Then HideToolTip: StartButtonOver = False
    If DesktopButtonOver = True Then HideToolTip: DesktopButtonOver = False
End Sub
