Attribute VB_Name = "Module1"
' Macro3 Macro
'
' Touche de raccourci du clavier: Ctrl+Shift+N
    
'The following statement

'Dim MyArray()

'declares an array without dimensions, so the compiler doesn't know how big it is and can't store anything inside of it.

'But you can use the ReDim statement to resize the array:

'ReDim MyArray(0 To 3)

'And if you need to resize the array while preserving its contents, you can use the Preserve keyword along with the ReDim statement:

'ReDim Preserve MyArray(0 To 3)


Sub AimportVac()



 
  Dim b(1 To 31, 1 To 2)
    For i = 1 To 31
        b(i, 1) = Worksheets("ccTDS").Cells(47, i).Value
    Next i
    
    For lig = LBound(b, 1) To UBound(b, 1)
        Cells(lig, 1) = b(lig, 1)
    Next lig
    
    ''' pour tests
    lvoArray = Array("N", "S2", "S1", "J2", "J1", "M", "M2", "M6", "J7", "J8", "J6", "J")
    
    
    Dim lvoArray2(1 To 12)
    For i = 1 To 12
        lvoArray2(i) = Worksheets("lvo").Cells(i + 1, 1).Value
    Next i
    
    ''' verif
    For lig = LBound(lvoArray2) To UBound(lvoArray2)
        Cells(lig, 9) = lvoArray2(lig)
    Next lig
        
         
    For lig = LBound(b, 1) To UBound(b, 1)
        If IsNumeric(Application.Match(b(lig, 1), lvoArray2, 0)) Then
        Cells(lig, 2) = b(lig, 1)
        Else: Cells(lig, 2) = "NIL"
        End If
    Next lig
    
    MsgBox "filtrage effectué"
    
    
End Sub
    
     
Sub FilterVac2()
    Dim lvo2(1 To 10, 1 To 2)
    For i = 2 To 6
        lvo2(i, 1) = Worksheets("lvo").Cells(i, 1).Value
        Cells(i, 1) = lvo2(i, 1)
       
    Next i
  
  End Sub
  
  
  
  
  
  
  
Sub ATab2D()
    
    
    
    i = 5
  Dim a(1 To 3, 1 To 2) ' 3 lignes x 2 colonnes
  a(1, 1) = "GRUE HOP"
  a(1, 2) = Range("I5").Value
  a(2, 1) = Worksheets("en cours").Range("A5").Value
  a(2, 2) = 22
  a(3, 1) = Worksheets("en cours").Range("A" & i).Hyperlinks(1).Address
  a(3, 2) = 32
  For lig = LBound(a, 1) To UBound(a, 1)
     For col = LBound(a, 2) To UBound(a, 2)
        Cells(lig, col) = a(lig, col)
     Next col
  Next lig
 
  Dim b(1 To 30, 1 To 2)
    For i = 5 To 15
        b(i, 1) = Worksheets("en cours").Range("A" & i).Value
       
    Next i
    
    For lig = LBound(b, 1) To UBound(b, 1)
        Cells(lig, 1) = b(lig, 1)
     
  Next lig
  
  End Sub
  
  
  Function CommentOf(incell) As String
' aceepts a cell as input and returns its comments (if any) back as a string "tour tdc. decontrole: Formation supports fh lyon"

On Error Resume Next
    Pattern = "tour tdc. decontrole:"
    CommentOf = incell.Comment.Text
    If InStr(CommentOf, Pattern) Then
        CommentOf = Right(CommentOf, Len(Pattern) + 1)
    End If
   
End Function

