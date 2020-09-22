Attribute VB_Name = "modFileAPI"
Private Const MAX_PATH = 260
Private Const SHGFI_DISPLAYNAME = &H200
Private Const SHGFI_EXETYPE = &H2000
Private Const SHGFI_SYSICONINDEX = &H4000  ' System icon index
Private Const SHGFI_LARGEICON = &H0        ' Large icon
Private Const SHGFI_SMALLICON = &H1        ' Small icon
Private Const ILD_TRANSPARENT = &H1        ' Display transparent
Private Const SHGFI_SHELLICONSIZE = &H4
Private Const SHGFI_TYPENAME = &H400
Private Const BASIC_SHGFI_FLAGS = SHGFI_TYPENAME _
   Or SHGFI_SHELLICONSIZE Or SHGFI_SYSICONINDEX _
   Or SHGFI_DISPLAYNAME Or SHGFI_EXETYPE

Private Type SHFILEINFO
    hIcon As Long
    iIcon As Long
    dwAttributes As Long
    szDisplayName As String * MAX_PATH
    szTypeName As String * 80
End Type

Private Declare Function SHGetFileInfo Lib "shell32.dll" Alias "SHGetFileInfoA" _
   (ByVal pszPath As String, _
    ByVal dwFileAttributes As Long, _
    psfi As SHFILEINFO, _
    ByVal cbSizeFileInfo As Long, _
    ByVal uFlags As Long) As Long

Private Declare Function ImageList_Draw Lib "comctl32.dll" _
   (ByVal himl&, ByVal i&, ByVal hDCDest&, _
    ByVal x&, ByVal y&, ByVal flags&) As Long

Private shinfo As SHFILEINFO
Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWND As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Public Enum IconSizes
    Icon_Small
    Icon_Large
End Enum

Public Sub DrawFileIcon(path As String, HDC As Long, Optional Size As IconSizes = Icon_Small, Optional x As Integer = 0, Optional y As Integer = 0)
    If Size = Icon_Small Then
        Icon = SHGetFileInfo(path, 0, shinfo, Len(shinfo), BASIC_SHGFI_FLAGS Or SHGFI_SMALLICON)
        ImageList_Draw Icon, shinfo.iIcon, HDC, x, y, ILD_TRANSPARENT
    Else
        Icon = SHGetFileInfo(path, 0, shinfo, Len(shinfo), BASIC_SHGFI_FLAGS Or SHGFI_LARGEICON)
        ImageList_Draw Icon, shinfo.iIcon, HDC, x, y, ILD_TRANSPARENT
    End If
End Sub

Public Sub ShellFile(path As String)
    ShellExecute frmTaskbar.hWND, "open", path, "", "", 1
End Sub

