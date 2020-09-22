Attribute VB_Name = "modDesktopIconPopup"
Public DesktopIconsShown As Boolean

Sub ToggleDesktopIcons()
If DesktopIconsShown Then
    DesktopIconsShown = False
    frmTaskbar.DesktopIconOff
    frmDesktopMenu.HideMenu
Else
    DesktopIconsShown = True
    frmTaskbar.DesktopIconOn
    frmDesktopMenu.DrawMenu "C:\windows\desktop\"
End If
End Sub

