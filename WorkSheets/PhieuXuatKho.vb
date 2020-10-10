Private Sub Worksheet_Change(ByVal Target As Range)
PXK.Unprotect
On Error Resume Next
If (Target.Value) <> "" Then
    If Target.Column = 3 Then
        If Target.Row >= 11 Then
            Call Ham_chung.dien_data_tu_dong(Target, PXK)
            Call Ham_chung.copy_sheet(PXK, Dieukien_MH, 11, "c", 1)
            Call Ton_kho_PXK.kho_co_de_xuat
        End If
    End If
End If
On Error GoTo 0
PXK.Protect
End Sub

Sub PXK_Luu()
Dim i As Long, lr_GHISO As Long, lr_PXK As Long
Dim thongbao_Xoa
Dim arr_N1, arr_N2
lr_GHISO = GHISO.Range("j" & Rows.Count).End(xlUp).Row
lr_PXK = PXK.Range("c" & Rows.Count).End(xlUp).Row
arr_N1 = PXK.Range("d2:j7")
arr_N2 = PXK.Range("c11:e" & lr_PXK)
arr_Dongia = PXK.Range("h11:h" & lr_PXK)
arr_Soluong = PXK.Range("g11:g" & lr_PXK)
arr_TTGC = PXK.Range("i11:j" & lr_PXK)
Call Tang_toc_code.tat_che_do

If Not Len(Trim(arr_N1(1, 7))) > 0 Then
    Application.Assistant.DoAlert "C" & ChrW(7843) & "nh Báo", "B" & ChrW(7841) & "n ch" & ChrW(432) & "a nh" & ChrW(7853) & "p s" & ChrW(7889) & " phi" & ChrW(7871) & "u", 0, 3, 0, 0, 1
ElseIf Not Len(Trim(arr_N1(4, 1))) > 0 Then
    Application.Assistant.DoAlert "C" & ChrW(7843) & "nh Báo", "B" & ChrW(7841) & "n ch" & ChrW(432) & "a nh" & ChrW(7853) & "p Ngày Xu" & ChrW(7845) & "t Kho", 0, 3, 0, 0, 0
    PXK.Range("d5").Select
ElseIf Application.WorksheetFunction.CountA(PXK.Range("c:c")) = 4 Then
    Application.Assistant.DoAlert "L" & ChrW(432) & "u ý", "C" & ChrW(7847) & "n ít nh" & ChrW(7845) & "t 1 Mã hàng", 0, 3, 0, 0, 1
ElseIf Application.WorksheetFunction.CountIf(GHISO.Range("e:e"), PXK.Range("i2")) > 0 Then
    Application.Assistant.DoAlert "C" & ChrW(7843) & "nh Báo", "S" & ChrW(7889) & " phi" & ChrW(7871) & "u này " & ChrW(273) & "ã t" & ChrW(7891) & "n t" & ChrW(7841) & "i", 0, 1, 0, 0, 1
ElseIf Application.WorksheetFunction.CountBlank(PXK.Range("c11:c" & lr_PXK)) > 0 Then
    Application.Assistant.DoAlert "C" & ChrW(7843) & "nh Báo", "Có ô tr" & ChrW(7889) & "ng trong d" & ChrW(7919) & " li" & ChrW(7879) & "u b" & ChrW(7841) & "n nh" & ChrW(7853) & "p", 0, 3, 0, 0, 0
    thongbao_Xoa = MsgBox("Ban có muôn xóa các ô trông không ?", vbYesNo + vbQuestion, "Thông Báo")
    If thongbao_Xoa = 6 Then
        Call Ham_chung.xoa_dongtrong(11, "h", 6, PXK)
    ElseIf thongbao_Xoa = 7 Then
        Call Tang_toc_code.bat_che_do
        Exit Sub
    End If
    
Else
    With GHISO
        Application.EnableEvents = False
        .Range("d" & lr_GHISO + 1).Resize(lr_PXK - 10, 1) = "XK"
        .Range("e" & lr_GHISO + 1).Resize(lr_PXK - 10, 1) = arr_N1(1, 7)
        .Range("f" & lr_GHISO + 1).Resize(lr_PXK - 10, 1) = arr_N1(4, 1)
        .Range("g" & lr_GHISO + 1).Resize(lr_PXK - 10, 1) = arr_N1(6, 1)
        .Range("h" & lr_GHISO + 1).Resize(lr_PXK - 10, 1) = arr_N1(6, 5)
        .Range("i" & lr_GHISO + 1).Resize(lr_PXK - 10, 1) = arr_N1(5, 1)
        .Range("j" & lr_GHISO + 1).Resize(lr_PXK - 10, 3) = arr_N2
        .Range("m" & lr_GHISO + 1).Resize(lr_PXK - 10, 1) = arr_Dongia
        .Range("n" & lr_GHISO + 1).Resize(lr_PXK - 10, 1) = arr_Soluong
        Application.EnableEvents = True ' Goi su kien change sheet GHISO
        .Range("o" & lr_GHISO + 1).Resize(lr_PXK - 10, 2) = arr_TTGC
    End With
    
    Application.Assistant.DoAlert "Thông Báo", "Xu" & ChrW(7845) & "t kho thành công", 0, 4, 0, 0, 1
End If

Call Tang_toc_code.bat_che_do
End Sub

Sub PXK_Taomoi()
Dim i As Byte, lr As Long

Application.EnableEvents = False
With PXK
    For i = 3 To 10
        lr = .Cells(Rows.Count, i).End(xlUp).Row
        If lr >= 11 Then
            .Unprotect
            .Range("d5") = Date
            .Range("d6:d7").ClearContents
            .Range("h5:h7").ClearContents
            .Range("c11:h" & lr).ClearContents
            .Range("j11:j" & lr).ClearContents
            .Protect
            Exit For
        End If
    Next
    Call Ham_chung.copy_sheet(PXK, Dieukien_MH, 11, "c", 1)
End With
Application.EnableEvents = True
End Sub
