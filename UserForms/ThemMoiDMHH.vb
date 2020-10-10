Private Sub cmd_Dong_Click()
Unload Themmoi_DMHH
End Sub
Private Sub cmd_Them_Click()
Dim i As Long, lr As Long
Dim arr_N
lr = DMHH.Range("C" & Rows.Count).End(xlUp).Row + 1
arr_N = DMHH.Range("C6:G" & lr - 1)

For i = 1 To UBound(arr_N, 1)
    If UCase(txt_MH) = UCase(arr_N(i, 1)) Then
        Application.Assistant.DoAlert "C" & ChrW(7843) & "nh Báo", "Mã hàng này " & ChrW(273) & "ã t" & ChrW(7891) & "n t" & ChrW(7841) & "i", 0, 1, 0, 0, 1
        Exit Sub
    End If
Next

With DMHH
    If Len(Trim(txt_DG)) > 0 And Len(Trim(txt_SL)) > 0 _
        And IsNumeric(txt_DG) And IsNumeric(txt_SL) Then
        .Range("C" & lr).Value = txt_MH
        .Range("D" & lr).Value = txt_TH
        .Range("E" & lr).Value = txt_DVT
        .Range("F" & lr).Value = CDbl(txt_DG)
        .Range("G" & lr).Value = CDbl(txt_SL)
        Application.Assistant.DoAlert "Thông Báo", "Hàng " & ChrW(273) & "ã " & ChrW(273) & ChrW(432) & ChrW(7907) & "c thêm vào", 0, 4, 0, 0, 1
    Else
        Application.Assistant.DoAlert "C" & ChrW(7843) & "nh Báo", "Thông tin cu" & ChrW(777) & "a ba" & ChrW(803) & "n không h" & ChrW(417) & ChrW(803) & "p lê" & ChrW(803), 0, 1, 0, 0, 1
    End If
End With

Set arr_N = Nothing
End Sub
Private Sub cmd_Xoa_Click()
Unload Themmoi_DMHH
Themmoi_DMHH.Show
End Sub

Private Sub txt_DG_Change()
txt_DG.Value = Format(txt_DG, "#,##0")
If Len(Trim(txt_DG)) > 0 And Len(Trim(txt_SL)) > 0 _
    And IsNumeric(txt_DG) And IsNumeric(txt_SL) Then
    txt_TT.Value = txt_DG * txt_SL
Else
    txt_TT.Value = ""
End If
End Sub

Private Sub txt_SL_Change()
If Len(Trim(txt_DG)) > 0 And Len(Trim(txt_SL)) > 0 _
    And IsNumeric(txt_DG) And IsNumeric(txt_SL) Then
    txt_TT.Value = txt_DG * txt_SL
Else
    txt_TT.Value = ""
End If
End Sub
Private Sub txt_TT_Change()
txt_TT.Value = Format(txt_TT, "#,##0")
End Sub
