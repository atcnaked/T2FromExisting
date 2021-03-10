Attribute VB_Name = "row_number_of_an_active_cell"


Sub rrrReturn_row_number_of_an_active_cell()
'declare a variable
Dim ws As Worksheet
Set ws = Worksheets("Analysis")

'get row number of an active cell and insert the value into cell A1
ws.Range("A1") = ActiveCell.Row

End Sub

Function Return_row_number_of_an_active_cell()

    Return_row_number_of_an_active_cell = ActiveCell.Row

End Function
