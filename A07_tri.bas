Attribute VB_Name = "A07_tri"
Sub Transpose()
 
  Dim b(1 To 31, 1 To 3)
    For i = 1 To 31
        b(i, 1) = Worksheets("choix agent+Mois+NomFichier").Cells(23, i + 1).Value
        b(i, 2) = Worksheets("choix agent+Mois+NomFichier").Cells(24, i + 1).Value
        b(i, 3) = Worksheets("choix agent+Mois+NomFichier").Cells(25, i + 1).Value
    Next i
        
    MsgBox "transpose effectué"
    
    indexLecture = 1
    indexEcriture = 1
    For indexLecture = 1 To 31
        
        If b(indexLecture, 3) = True Then
            'MsgBox "Match !"
            b(indexEcriture, 1) = b(indexLecture, 1)
            b(indexEcriture, 2) = b(indexLecture, 2)
            indexEcriture = indexEcriture + 1
        End If
    Next indexLecture
    

    MsgBox "tri effectué"
    
    
    
    
    
    'affichage
    For lig = 1 To indexEcriture - 1
        Cells(lig, 1) = b(lig, 1)
        Cells(lig, 2) = b(lig, 2)
       ' Cells(lig, 3) = b(lig, 3)
    Next lig
    
End Sub

Sub tri()
  ' il faut compter le nombre de jour avant de dimensionner le tableau
  
  
  Dim vacList(1 To 12)
    For i = 1 To 12
        vacList(i) = Worksheets("VAC").Cells(i + 1, 1).Value
    Next i
    
  
  indexLecture = 1
  indexEcriture = 1
  Dim b(1 To 31, 1 To 2)
    For indexLecture = 1 To 31
        
        If IsNumeric(Application.Match(b(indexLecture, 2), vacList, 0)) Then
            MsgBox "Match !"
            b(indexEcriture, 1) = b(indexLecture, 1)
            b(indexEcriture, 2) = b(indexLecture, 2)
            indexEcriture = indexEcriture + 1
        End If
    Next indexLecture
    

    MsgBox "tri effectué"
    
    
End Sub









Sub copieAndTransposeTDS()

    Range("A1:AK122").Select
    Selection.Copy
    Sheets("transposed TDS").Select
    Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=True
End Sub
