Attribute VB_Name = "modCustomApplets"
Public ToolTipDisplayed As Boolean

Public Sub MsgBoxEX()

End Sub

Public Sub ToolTipEX(Text As String, Top, Left)
    If ToolTipDisplayed = False Then
        frmToolTip.DisplayTip Text, Val(Top), Val(Left)
    End If
End Sub

Public Sub HideToolTip()
    frmToolTip.HideTip
End Sub
