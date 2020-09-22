Attribute VB_Name = "modScreenData"
Public Function ScreenWidth() As Integer
    ScreenWidth = Screen.Width / Screen.TwipsPerPixelX
End Function

Public Function ScreenHeight() As Integer
    ScreenHeight = Screen.Height / Screen.TwipsPerPixelY
End Function

