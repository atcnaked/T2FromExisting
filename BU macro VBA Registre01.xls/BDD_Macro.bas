Attribute VB_Name = "BDD_Macro"
Sub suppronglet()
' suppress lat tab

totalSheetNumber = Worksheets.Count

While totalSheetNumber > 6
    
    Worksheets(totalSheetNumber).Select
    ActiveWindow.SelectedSheets.Delete
    totalSheetNumber = Worksheets.Count
Wend
    

    
End Sub

Sub pdf2()
'
' pdf2 Macro
'

'
    Range("A1:I29").Select
    Range("I29").Activate
    Selection.PrintOut Copies:=1, Collate:=True
End Sub

Sub pdf3() ' non testé

 If IsFileOpen(Sheets("1").Cells(1, 3) & "%") Then
    Else
        Ligne = 5                                                      'Ligne de départ
            While Cells(Ligne, 2) <> ""                       'Recherche la première cellule vide
                Ligne = Ligne + 1
            Wend

        Ligne = Ligne + 4

        Range(Cells(1, 1), Cells(Ligne, 12)).Select     'Selectionne la plage de cellule

            With Worksheets("***")
                Selection.ExportAsFixedFormat _
                Type:=xlTypePDF, _
                Filename:=Sheets("1").Cells(1, 3) & "%", _
                Quality:=xlQualityStandard, _
                IncludeDocProperties:=True, _
                IgnorePrintAreas:=True, _
                OpenAfterPublish:=False
            End With

    End If
End Sub



