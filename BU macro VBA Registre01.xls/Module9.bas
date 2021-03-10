Attribute VB_Name = "Module9"
Public Sub ActiveRow()
 MsgBox ActiveCell.Row
End Sub

Public Sub addArow()
 ' MsgBox ActiveCell.Row
 'Color = Range("A" & ActiveCell.Row).Interior.Color
 'Range("A" & ActiveCell.Row).Interior.Color = 0
 Range("A" & ActiveCell.Row) = Range("A" & ActiveCell.Row) + 1
 Range("A" & ActiveCell.Row).Interior.Color = 5
End Sub

