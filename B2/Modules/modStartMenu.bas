Attribute VB_Name = "modStartMenu"
Public StartMenuShown As Boolean

Sub ToggleStartMenu()
If StartMenuShown Then
    StartMenuShown = False
    frmTaskbar.StartButtonOff
    frmStartMenu.HideMe
Else
    StartMenuShown = True
    frmTaskbar.StartButtonOn
    frmStartMenu.Display
End If
End Sub
