Attribute VB_Name = "Module6"
Sub onglet_affiche_suppr_hide()
Attribute onglet_affiche_suppr_hide.VB_ProcData.VB_Invoke_Func = " \n14"
'
' onglet_affiche_suppr_hide Macro
'

'
    Sheets.Add After:=ActiveSheet
    Sheets("Feuil2").Select
    Sheets("Feuil2").Name = "myonglet"
    Sheets("Feuil1").Select
    ActiveWindow.SelectedSheets.Delete
    Sheets("myonglet").Select
    ActiveWindow.SelectedSheets.Visible = False
End Sub
