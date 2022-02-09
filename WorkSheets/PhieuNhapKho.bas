Private Sub Worksheet_Change(ByVal Target As Range)
On Error Resume Next
If (Target.Value) <> "" Then Call Ham_chung.dien_data_tu_dong(Target, PNK)
On Error GoTo 0
End Sub

Sub PNK_Luu()
Dim i As Long, lr_GHISO As Long, lr_PNK As Long
Dim thongbao_Xoa
Dim arr_N1, arr_N2
lr_GHISO = GHISO.Range("j" & Rows.Count).End(xlUp).Row
lr_PNK = PNK.Range("c" & Rows.Count).End(xlUp).Row
arr_N1 = PNK.Range("d2:i7")
arr_N2 = PNK.Range("c11:i" & lr_PNK)
Call Tang_toc_code.tat_che_do

If Not Len(Trim(arr_N1(1, 6))) > 0 Then
    Application.Assistant.DoAlert "C" & ChrW(7843) & "nh Báo", "B" & ChrW(7841) & "n ch" & ChrW(432) & "a nh" & ChrW(7853) & "p s" & ChrW(7889) & " phi" & ChrW(7871) & "u", 0, 3, 0, 0, 1
ElseIf Not Len(Trim(arr_N1(4, 1))) > 0 Then
    Application.Assistant.DoAlert "C" & ChrW(7843) & "nh Báo", "B" & ChrW(7841) & "n ch" & ChrW(432) & "a nh" & ChrW(7853) & "p Ngày Nh" & ChrW(7853) & "p Kho", 0, 3, 0, 0, 0
    PNK.Range("d5").Select
ElseIf Application.WorksheetFunction.CountA(PNK.Range("c:c")) = 4 Then
    Application.Assistant.DoAlert "L" & ChrW(432) & "u ý", "C" & ChrW(7847) & "n ít nh" & ChrW(7845) & "t 1 Mã hàng", 0, 3, 0, 0, 1
ElseIf Application.WorksheetFunction.CountIf(GHISO.Range("e:e"), PNK.Range("i2")) > 0 Then
    Application.Assistant.DoAlert "C" & ChrW(7843) & "nh Báo", "S" & ChrW(7889) & " phi" & ChrW(7871) & "u này " & ChrW(273) & "ã t" & ChrW(7891) & "n t" & ChrW(7841) & "i", 0, 1, 0, 0, 1
ElseIf Application.WorksheetFunction.CountBlank(PNK.Range("c11:c" & lr_PNK)) > 0 Then
    Application.Assistant.DoAlert "C" & ChrW(7843) & "nh Báo", "Có ô tr" & ChrW(7889) & "ng trong d" & ChrW(7919) & " li" & ChrW(7879) & "u b" & ChrW(7841) & "n nh" & ChrW(7853) & "p", 0, 3, 0, 0, 0
    thongbao_Xoa = MsgBox("Ban có muôn xóa các ô trông không ?", vbYesNo + vbQuestion, "Thông Báo")
    If thongbao_Xoa = 6 Then
        Call Ham_chung.xoa_dongtrong(11, "g", 5, PNK)
    ElseIf thongbao_Xoa = 7 Then
        Call Tang_toc_code.bat_che_do
        Exit Sub
    End If
    
Else
    With GHISO
        Application.EnableEvents = False
        .Range("d" & lr_GHISO + 1).Resize(lr_PNK - 10, 1) = "NK"
        .Range("e" & lr_GHISO + 1).Resize(lr_PNK - 10, 1) = arr_N1(1, 6)
        .Range("f" & lr_GHISO + 1).Resize(lr_PNK - 10, 1) = arr_N1(4, 1)
        .Range("g" & lr_GHISO + 1).Resize(lr_PNK - 10, 1) = arr_N1(6, 1)
        .Range("h" & lr_GHISO + 1).Resize(lr_PNK - 10, 1) = arr_N1(6, 4)
        .Range("i" & lr_GHISO + 1).Resize(lr_PNK - 10, 1) = arr_N1(5, 1)
        Application.EnableEvents = True ' Goi su kien change sheet GHISO
        .Range("j" & lr_GHISO + 1).Resize(lr_PNK - 10, 7) = arr_N2
    End With
    
    Application.Assistant.DoAlert "Thông Báo", "Nh" & ChrW(7853) & "p kho thành công", 0, 4, 0, 0, 1
End If

Call Tang_toc_code.bat_che_do
End Sub

Sub PNK_Taomoi()
Dim i As Byte, lr As Long

Application.EnableEvents = False
With PNK
    For i = 3 To 9
        lr = .Cells(Rows.Count, i).End(xlUp).Row
        If lr >= 11 Then
            .Unprotect
            .Range("d5") = Date
            .Range("d6:d7").ClearContents
            .Range("g7").ClearContents
            .Range("c11:g" & lr).ClearContents
            .Range("i11:i" & lr).ClearContents
            .Protect
            Exit For
        End If
    Next
End With
Application.EnableEvents = True
End Sub
