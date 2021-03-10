Attribute VB_Name = "Module4"
Sub PDFprint()
Attribute PDFprint.VB_ProcData.VB_Invoke_Func = " \n14"
'
' PDFprint Macro
'

'
    Range("A1:I29").Select
    ' Range("I29").Activate
    ChDir "C:\Users\tourdecontrole\Desktop"
    'ActiveSheet.ExportAsFixedFormat _ ligne d origine
    Selection.ExportAsFixedFormat _
        Type:=xlTypePDF, _
        Filename:="C:\Users\tourdecontrole\Desktop\VBA Registre06" & Range("R17").Value & ".pdf", _
        Quality:=xlQualityStandard, _
        IncludeDocProperties:=True, _
        IgnorePrintAreas:=True, _
        OpenAfterPublish:=True
End Sub
