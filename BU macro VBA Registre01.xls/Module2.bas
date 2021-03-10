Attribute VB_Name = "Module2"

Function IsWorkBookOpenF(Name As String) As Boolean
    Dim xWb As Workbook
    On Error Resume Next
    Set xWb = Application.Workbooks.Item(Name)
    IsWorkBookOpenF = (Not xWb Is Nothing)
End Function

 
Sub IsWorkBookOpen()
    Dim xRet As Boolean
    xRet = IsWorkBookOpenF("TDS 2021.xlsx")
    If xRet Then
        MsgBox "The file is open, carry on", vbInformation, "Kutools for Excel"
    Else
        MsgBox "TDS 2021.xlsx is not open: open it please", vbInformation, "Kutools for Excel"
    End If
End Sub


Sub AbetaimportVac()
Attribute AbetaimportVac.VB_ProcData.VB_Invoke_Func = "M\n14"
    MsgBox "pour éviter les soucis penser à renomer TDS 2021.xlsx en TDS 2021TEST.xlsx"
    Windows("TDS 2021TEST.xlsx").Activate
    Sheets("Janvier").Select
 
 controlerLine = 47
  Dim b(1 To 32, 1 To 4)
    For i = 1 To 32
        b(i, 1) = Cells(1, i).Value
        For k = 0 To 1
        b(i, 2) = Cells(controlerLine, i).Value
        b(i, 3) = Cells(controlerLine + 1, i).Value
        b(i, 4) = CommentOf(Cells(controlerLine + 1, i))
        Next k
    Next i
    
    Windows("VBA Registre01.xls").Activate
    Sheets.Add After:=ActiveSheet
    
    For lig = LBound(b, 1) To UBound(b, 1)
        Cells(lig, 1) = b(lig, 1)
        Cells(lig, 2) = b(lig, 2)
        Cells(lig, 3) = b(lig, 3)
        Cells(lig, 4) = b(lig, 4)
    Next lig
    
    Cells(1, 3) = "ligne 2"
    Cells(1, 4) = "commentaires"
    
    MsgBox "copie effectué"
    
End Sub



  Function CommentOf(incell) As String
    ' aceepts a cell as input and returns its comments (if any) back as a string "tour tdc. decontrole: Formation supports fh lyon"

    On Error Resume Next
    ' test Pattern = "tour tdc. decontrole:"
    'Pattern = ""
    CommentOf = incell.Comment.Text
    'If InStr(CommentOf, Pattern) Then
    '    CommentOf = Right(CommentOf, Len(Pattern) + 1)
    'End If
End Function

