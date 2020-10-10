Private Sub cb_MH_Change()
Dim i As Long, lr As Long
Dim arr_N
lr = DMHH.Range("C" & Rows.Count).End(xlUp).Row + 1
arr_N = DMHH.Range("C6:G" & lr - 1)

For i = 1 To UBound(arr_N, 1)
    If UCase(cb_MH) = UCase(arr_N(i, 1)) Then
        txt_TH = arr_N(i, 2)
        txt_DVT = arr_N(i, 3)
        txt_DG = arr_N(i, 4)
        txt_SL = arr_N(i, 5)
    End If
Next

Set arr_N = Nothing
End Sub

Private Sub cmd_Capnhat_Click()
Dim i As Long, lr As Long
Dim arr_N
lr = DMHH.Range("C" & Rows.Count).End(xlUp).Row
arr_N = DMHH.Range("C6:G" & lr)

With DMHH
    If Len(Trim(txt_DG)) > 0 And Len(Trim(txt_SL)) > 0 _
        And IsNumeric(txt_DG) And IsNumeric(txt_SL) Then
        
        For i = 1 To UBound(arr_N, 1)
            If UCase(cb_MH) = UCase(arr_N(i, 1)) Then
                arr_N(i, 2) = txt_TH
                arr_N(i, 3) = txt_DVT
                arr_N(i, 4) = txt_DG
                arr_N(i, 5) = txt_SL
            End If
        Next
    Else
        Application.Assistant.DoAlert "C" & ChrW(7843) & "nh Báo", "Thông tin cu" & ChrW(777) & "a ba" & ChrW(803) & "n không h" & ChrW(417) & ChrW(803) & "p lê" & ChrW(803), 0, 1, 0, 0, 1
    End If
End With

DMHH.Range("C6").Resize(lr - 5, 5) = arr_N
Set arr_N = Nothing
End Sub

Private Sub cmd_HUYBO_Click()
Unload Chinhsua_DMHH
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

Private Sub UserForm_Initialize()
Dim mh
Call Tang_toc_code.tat_che_do
For Each mh In [Ma_hang] 'Nap cac ma hang trong name dong vao combobox
    cb_MH.AddItem mh
Next
Call Tang_toc_code.bat_che_do
End Sub
