Attribute VB_Name = "modGraphicAPICalls"
Declare Function StretchBlt Lib "gdi32" (ByVal HDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long

Declare Function DrawIconEx Lib "user32" (ByVal HDC As Long, ByVal xLeft As Long, ByVal yTop As Long, ByVal hIcon As Long, ByVal cxWidth As Long, ByVal cyWidth As Long, ByVal istepIfAniCur As Long, ByVal hbrFlickerFreeDraw As Long, ByVal diFlags As Long) As Long
'Constants for DrawIconEX
Public Const DI_NORMAL = &H3
'Functions for GetIcon
Private Declare Function GetClassLong Lib "user32" Alias "GetClassLongA" (ByVal hWND As Long, ByVal nIndex As Integer) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWND As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function SendMessageTimeout Lib "user32" Alias "SendMessageTimeoutA" (ByVal hWND As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long, ByVal fuFlags As Long, ByVal uTimeout As Long, lpdwResult As Long) As Long
'Contstants for GetIcon
Private Const WM_GETICON = &H7F
Private Const GCL_HICON = (-14)
Private Const GCL_HICONSM = (-34)
Private Const WM_QUERYDRAGICON = &H37

Declare Function AlphaBlending Lib "Alphablending.dll" _
                     (ByVal destHDC As Long, ByVal XDest As Long, ByVal YDest As Long, _
                      ByVal destWidth As Long, ByVal destHeight As Long, ByVal srcHDC As Long, _
                      ByVal xSrc As Long, ByVal ySrc As Long, ByVal srcWidth As Long, ByVal srcHeight As Long, ByVal AlphaSource As Long) As Long

Declare Function GetDC Lib "user32" (ByVal hWND As Long) As Long
Declare Function ReleaseDC Lib "user32" (ByVal hWND As Long, ByVal HDC As Long) As Long


Public Function GetIcon(hWND As Long) As Long
    Call SendMessageTimeout(hWND, WM_GETICON, 0, 0, 0, 1000, GetIcon)
    If Not CBool(GetIcon) Then GetIcon = GetClassLong(hWND, GCL_HICONSM)
    If Not CBool(GetIcon) Then Call SendMessageTimeout(hWND, WM_GETICON, 1, 0, 0, 1000, GetIcon)
    If Not CBool(GetIcon) Then GetIcon = GetClassLong(hWND, GCL_HICON)
    If Not CBool(GetIcon) Then Call SendMessageTimeout(hWND, WM_QUERYDRAGICON, 0, 0, 0, 1000, GetIcon)
End Function

Sub BltDesktop(sourceX As Integer, sourceY As Integer, targetBox As PictureBox)
    deskHDC = GetDC(0)
    BitBlt targetBox.HDC, 0, 0, targetBox.Width, targetBox.Height, deskHDC, sourceX, sourceY, vbSrcCopy
    ReleaseDC 0, deskHDC
End Sub
