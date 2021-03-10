Attribute VB_Name = "A0z_open_close_file"
Sub openfile()
Attribute openfile.VB_ProcData.VB_Invoke_Func = " \n14"
'
' openfile Macro
'

'
    Workbooks.Open Filename:= _
        "C:\A Trav0514 2021\2021-01 VBA Excel Tri dynamique\TDS 2021.xlsx"
    Sheets("Février").Select
    Cells.Select
    Selection.Copy
    Windows("VBA Registre01.xls").Activate
    Cells.Select
    ActiveSheet.Paste
End Sub
Sub open_copytab_closefile()
Attribute open_copytab_closefile.VB_ProcData.VB_Invoke_Func = " \n14"
'
' closefile Macro
'

'
    Workbooks.Open Filename:= _
        "C:\A Trav0514 2021\2021-01 VBA Excel Tri dynamique\TDS 2021.xlsx"
    Cells.Select
    Selection.Copy
    Windows("VBA Registre01.xls").Activate
    Windows("TDS 2021.xlsx").Activate
    Windows("VBA Registre01.xls").Activate
    Sheets.Add After:=ActiveSheet
    ActiveSheet.Paste
    Windows("TDS 2021.xlsx").Activate
    ActiveWindow.Close
End Sub
Sub tdsselecttab()
Attribute tdsselecttab.VB_ProcData.VB_Invoke_Func = " \n14"
'
' tdsselecttab Macro
'

'
    Workbooks.Open Filename:= _
        "C:\A Trav0514 2021\2021-01 VBA Excel Tri dynamique\TDS 2021.xlsx"
    Sheets("Février").Select
    Cells.Select
    Selection.Copy
    Sheets("Janvier").Select
    Cells.Select
    Application.CutCopyMode = False
    Selection.Copy
    Windows("VBA Registre01.xls").Activate
    Sheets.Add After:=ActiveSheet
    Cells.Select
    ActiveSheet.Paste
    Windows("TDS 2021.xlsx").Activate
    ActiveWindow.Close
End Sub

' tvariableglobale Macro
Sub tvariableglobale()
Attribute tvariableglobale.VB_ProcData.VB_Invoke_Func = "M\n14"
    ' the name's cell is "bato" !
    MsgBox (Range("bato").Value) '
    Range("bato").Value = Range("bato").Value + 1
    MsgBox (Range("bato").Value)


End Sub

2
3
4
5
6
7
8
9
10
11
12
13
14
15
16
17
    

