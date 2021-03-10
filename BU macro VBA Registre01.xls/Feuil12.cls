VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Feuil12"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    ' très rapide sur un tableau plus léger
    
    formerLine = Range("A1").Value
    Cells(formerLine, 4).Value = ""
    Cells(formerLine, 5).Font.Color = 0
    Range("A1").Value = ActiveCell.Row
    Cells(ActiveCell.Row, 4).Value = "-->"
    Cells(ActiveCell.Row, 5).Font.Color = 16000000
    
    'MsgBox ("debug !")
End Sub
