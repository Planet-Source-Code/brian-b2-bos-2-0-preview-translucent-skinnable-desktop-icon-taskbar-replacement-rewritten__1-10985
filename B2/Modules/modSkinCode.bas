Attribute VB_Name = "modSkinCode"
Function GetSkinName() As String
    GetSkinName = GetSetting("B2", "Skin", "SkinName", "B2 Default")
End Function

Function GetTaskbarButtonWidth() As Integer 'Loads the width of the taskbar program button
    GetTaskbarButtonWidth = Val(GetSetting("B2", "Skin", "TaskbarButtonWidth", "200"))
End Function

Function GetTaskbarHeight() As Integer 'Loads the height of the taskbar
    GetTaskbarHeight = Val(GetSetting("B2", "Skin", "TaskbarHeight", "30"))
End Function

Function GetStartButtonWidth() As Integer 'Loads the width of the start button
    GetStartButtonWidth = Val(GetSetting("B2", "Skin", "StartButtonWidth", "30"))
End Function

Function GetDesktopButtonWidth()
    GetDesktopButtonWidth = Val(GetSetting("B2", "Skin", "DesktopButtonWidth", "30"))
End Function

Function GetSkinImage(name As String) As String
    GetSkinImage = App.Path & "\skins\" & GetSkinName & "\" & name
End Function

Function GetStartMenuWidth() As Integer
    GetStartMenuWidth = Val(GetSetting("B2", "Skin", "StartMenuWidth", "220"))
End Function

Function GetStartMenuHeight() As Integer
    GetStartMenuHeight = Val(GetSetting("B2", "Skin", "StartMenuHeight", "320"))
End Function
