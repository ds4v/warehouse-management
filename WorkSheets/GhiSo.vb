Private Sub Worksheet_Change(ByVal Target As Range)
Call Tang_toc_code.tat_che_do

On Error Resume Next
If (Target.Value) <> "" Then
    If Target.Row >= 6 Then
        If Target.Column >= 3 Then
            Call Ham_chung.copy_sheet(GHISO, copy_GHISO, 6, "p", 14)
            Call Ton_kho_PXK.kho_co_de_xuat
        End If
    End If
End If
On Error GoTo 0

Call Tang_toc_code.bat_che_do
End Sub

Sub GHISO_Taomoi()
Dim lr_GHISO As Long
lr_GHISO = GHISO.Range("D" & Rows.Count).End(xlUp).Row + 1

GHISO.Unprotect
GHISO.Range("C6:P" & lr_GHISO).ClearContents
GHISO.Protect
End Sub
