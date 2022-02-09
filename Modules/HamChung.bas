Sub dien_data_tu_dong(Target As Range, sources As Worksheet)
Call Tang_toc_code.tat_che_do
Dim socot As Byte
socot = IIf(sources.CodeName = "PNK", 5, 6)
If Target.Column = 3 Then
    If Target.Row >= 11 Then ' Vung cho phep thao tac
        On Error Resume Next
        If Len(Trim(Target.Offset(-1, 0))) > 0 Then ' Kiem tra ô trông
            sources.Unprotect
            Call Ham_chung.dotim_theo_mahang(Target, socot)
        Else
            Application.Assistant.DoAlert "C" & ChrW(7843) & "nh Báo", "Có ô tr" & ChrW(7889) & "ng phía trên", 0, 3, 0, 0, 1
            Application.EnableEvents = False
            Target.Value = Empty ' Neu co ô trông phia tren thi se xoa dong do
            Application.EnableEvents = True
        End If
            sources.Protect
        On Error GoTo 0
    End If
End If
Call Tang_toc_code.bat_che_do
End Sub

Sub copy_sheet(sources As Worksheet, sheet_paste As Worksheet, tu_hang As Long, den_cot As String, socot As Long)
Dim lr As Long
lr = sources.Range("D" & Rows.Count).End(xlUp).Row
arr_N = sources.Range("C" & tu_hang & ":" & den_cot & lr)
sheet_paste.Range("A2").Resize(1000000, socot).Clear

If lr <> tu_hang Then
    sheet_paste.Range("A2").Resize(UBound(arr_N, 1), socot) = arr_N
Else
    sheet_paste.Range("A2").Resize(1, socot) = arr_N
End If
End Sub

Sub dotim_theo_mahang(Ma_hang, socot As Byte)
Dim i As Long, j As Long, lr As Long
Dim arr_N, arr_Dulieu
lr = DMHH.Range("c" & Rows.Count).End(xlUp).Row
arr_N = DMHH.Range("c6:g" & lr)
ReDim arr_Dulieu(1 To 1, 1 To socot)
For i = 1 To lr - 5
    If Ma_hang = arr_N(i, 1) Then
        arr_Dulieu(1, 1) = arr_N(i, 1)
        arr_Dulieu(1, 2) = arr_N(i, 2)
        arr_Dulieu(1, 3) = arr_N(i, 3)
        arr_Dulieu(1, 4) = arr_N(i, 4)
        arr_Dulieu(1, 5) = arr_N(i, 5)
        If socot > 5 Then
            arr_Dulieu(1, socot) = arr_N(i, 4)
            arr_Dulieu(1, socot - 1) = arr_N(i, 5) 'Ma_hang.offset(0,4)
            arr_Dulieu(1, socot - 2) = arr_N(i, 5)
        End If
        Exit For
    End If
Next

Application.EnableEvents = False
Ma_hang.Resize(, socot) = arr_Dulieu
Set arr_Dulieu = Nothing
Application.EnableEvents = True
End Sub

Sub xoa_dongtrong(Vitri As Long, cotcuoi As String, socot As Long, Sh As Worksheet)
Dim i As Long, k As Long, lr As Long
Dim arr_N, arr_D
lr = Sh.Range("C" & Rows.Count).End(xlUp).Row
arr_N = Sh.Range("C" & Vitri & ":" & cotcuoi & lr)
ReDim arr_D(1 To UBound(arr_N, 1), 1 To 6)
k = 0

Application.ScreenUpdating = False
For i = 1 To UBound(arr_N, 1)
    If Len(Trim(arr_N(i, 1))) > 0 Then
        k = k + 1
        arr_D(k, 1) = arr_N(i, 1)
        arr_D(k, 2) = arr_N(i, 2)
        arr_D(k, 3) = arr_N(i, 3)
        arr_D(k, 4) = arr_N(i, 4)
        arr_D(k, 5) = arr_N(i, 5)
        arr_D(k, 6) = arr_N(i, socot)
    End If
Next

Application.EnableEvents = False
Sh.Unprotect
Sh.Range("c" & Vitri).Resize(UBound(arr_N, 1), socot) = arr_D
Sh.Protect
Application.EnableEvents = True

Set arr_N = Nothing
Set arr_D = Nothing
Application.ScreenUpdating = True
End Sub

Public Function TransposeArray(myarray As Variant) As Variant
Dim X As Long
Dim Y As Long
Dim Xupper As Long
Dim Yupper As Long
Dim tempArray As Variant
    Xupper = UBound(myarray, 2)
    Yupper = UBound(myarray, 1)
    ReDim tempArray(Xupper, Yupper)
    For X = 0 To Xupper
        For Y = 0 To Yupper
            tempArray(X, Y) = myarray(Y, X)
        Next Y
    Next X
    TransposeArray = tempArray
End Function
