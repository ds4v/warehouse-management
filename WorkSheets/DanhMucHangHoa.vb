Private Sub Worksheet_Change(ByVal Target As Range)

On Error Resume Next 'De khi keo du lieu xuong va nhân bua ko bi loi
If Target.Row < 7 Then Exit Sub 'Neu o hien tai trong thi ko co gi xay ra

Call Tang_toc_code.tat_che_do
Call Ham_chung.copy_sheet(DMHH, copy_DMHH, 6, "h", 7)

'''''''' Dien so thu tu de sap xep du lieu khi dua vao sheet THNXT '''''''
Dim arr_STT, i As Long, lr_copy_DMHH As Long
lr_copy_DMHH = copy_DMHH.Range("D" & Rows.Count).End(xlUp).Row - 1
If lr_copy_DMHH > 0 Then
    ReDim arr_STT(1 To lr_copy_DMHH, 1 To 1)
    For i = 1 To lr_copy_DMHH
        arr_STT(i, 1) = i
    Next
    copy_DMHH.Range("G2").Resize(UBound(arr_STT), 1) = arr_STT
End If
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim k As Long, lr As Long
Dim arr_N, Ma_hang
lr = Target.Row - 1
arr_N = DMHH.Range("C6:G" & lr)
'Nap du lieu tu ô dau den ô truoc ô vua nhap, de luc quet mang ko quet trung ô vua nhap, neu ko no se nghi ô vua nhap bi trung

Ma_hang = Target.Value
'Quet mang, neu trung thi thông bao va xoa
For i = 1 To UBound(arr_N, 1)
    If UCase(Ma_hang) = UCase(arr_N(i, 1)) Then
        Application.EnableEvents = False ' Vô hieu hoa su kien de khi xoa dong ko bi anh huong
        Application.Assistant.DoAlert "C" & ChrW(7843) & "nh Báo", "Mã hàng " & Ma_hang & " " & ChrW(273) & "ã t" & ChrW(7891) & "n t" & ChrW(7841) & "i", 0, 1, 0, 0, 1
        Call Tang_toc_code.bat_che_do
        On Error GoTo 0
        Set arr_N = Nothing
        Exit Sub
    End If
Next

On Error GoTo 0
Set arr_N = Nothing

Call Ham_chung.copy_sheet(DMHH, copy_DMHH, 6, "h", 6)
Call Tang_toc_code.bat_che_do
End Sub

Sub xoa_dongtrong_DMHH()
Call Ham_chung.xoa_dongtrong(6, "g", 5, DMHH)
Call Ham_chung.copy_sheet(DMHH, copy_DMHH, 6, "h", 6)
End Sub

Sub tim_kiem_DMHH()
Dim strsearch As String
strsearch = LCase([H1])
Dim lr As Long
lr = DMHH.Range("c" & Rows.Count).End(xlUp).Row
Dim rw As Range, r As Range
Set r = DMHH.Range("c6:h" & lr)

For Each rw In r.Rows
    If InStr(LCase(DMHH.Cells(rw.Row, 3)), strsearch) Then
        rw.Select
        Exit Sub
    End If
Next rw
End Sub
