Attribute VB_Name = "Module8"
Sub explicationcouleur()
Attribute explicationcouleur.VB_ProcData.VB_Invoke_Func = " \n14"
'
' explicationcouleur Macro
'

'
    
    
    Range("F5:G6").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 65535
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    MsgBox ("sélectionne ton nom mois et année là")
    
    Range("F9:G10").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 65535
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    Range("F5:G6").Select
    With Selection.Interior
        .Pattern = xlNone
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    MsgBox ("et appuie ici p générer !")
End Sub
Sub insertbouton()
Attribute insertbouton.VB_ProcData.VB_Invoke_Func = " \n14"
'
' insertbouton Macro
'

'
    Range("M14").Select
    ActiveSheet.OptionButtons.Add(740.25, 119.25, 72, 72).Select
    Range("L12").Select
End Sub

' Excellent !! crée des boutons leur associe une fonction !!!
Sub a()
Attribute a.VB_ProcData.VB_Invoke_Func = "H\n14"
  Dim btn As Button
  Application.ScreenUpdating = False
  ActiveSheet.Buttons.Delete
  Dim t As Range
  For i = 2 To 6 Step 2
    Set t = ActiveSheet.Range(Cells(i, 3), Cells(i, 3))
    Set btn = ActiveSheet.Buttons.Add(t.Left, t.Top, t.Width, t.Height)
    With btn
      .OnAction = "btnS"
      .Caption = "BtnCaption " & i
      .Name = "BtnName" & i
    End With
  Next i
  Application.ScreenUpdating = True
End Sub

Sub btnS()
 MsgBox Application.Caller
End Sub

