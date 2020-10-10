Sub kho_co_de_xuat()
Dim cn As Object, rs As Object, Dulieu As String, lr_GHISO As Long
Dim lr_VLOOKUP_MH As Long, lr_SUMIFS_NK As Long, lr_SUMIFS_XK As Long
Dim lr_Dieukien_MH As Long
Dim i As Long, j As Long, k As Long
Dim arr_VLOOKUP_MH, arr_SUMIFS_NK, arr_SUMIFS_XK, arr_D
Dim arr_theo_MH, nap_arr_theo_MH
Dim Dic As Object

Call Tang_toc_code.tat_che_do

' Lay ra cac mang du lieu trong cong thuc
' IFERROR(VLOOKUP(C11,DMHH,5,0)+SUMIFS(GHISO!N:N,GHISO!D:D,"NK",GHISO!J:J,C11)-SUMIFS(GHISO!N:N,GHISO!D:D,"XK",GHISO!J:J,C11),"")

lr_GHISO = GHISO.Range("D" & Rows.Count).End(xlUp).Row
Ton_trong_kho.UsedRange.Clear

Set cn = CreateObject("ADODB.Connection")
Set rs = CreateObject("ADODB.Recordset")

cn.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & Chr(39) & _
ThisWorkbook.FullName & Chr(39) & ";Extended Properties=""Excel 12.0 Xml;HDR=YES;IMEX=1"";"
cn.Open

Du_lieu_MH = "(SELECT ma_hang FROM [EXCEL 12.0;Database=" & ThisWorkbook.FullName & "].[Dieukien_MH$])"

' Cac ma hang duoc chon va so luong ban dau
Set rs = cn.Execute("SELECT ma_hang, so_luong FROM [copy_DMHH$] WHERE ma_hang IN " & Du_lieu_MH)
Ton_trong_kho.Range("A2").CopyFromRecordset rs ' VLOOKUP(C11, DMHH, 5, 0)

If lr_GHISO <> 5 Then
    ' Cac ma hang da nhap kho va so luong dc nhap
    Set rs = cn.Execute("SELECT ma_hang,SUM(so_luong) FROM [copy_GHISO$] WHERE loai ='NK' GROUP BY ma_hang")
    Ton_trong_kho.Range("C2").CopyFromRecordset rs ' SUMIFS(GHISO!N:N,GHISO!D:D,"NK",GHISO!J:J,C11)
    
    ' Cac ma hang da xuat kho va so luong dc xuat
    Set rs = cn.Execute("SELECT ma_hang,SUM(so_luong) FROM [copy_GHISO$] WHERE loai ='XK' GROUP BY ma_hang")
    Ton_trong_kho.Range("E2").CopyFromRecordset rs ' SUMIFS(GHISO!N:N,GHISO!D:D,"XK",GHISO!J:J,C11)
    
    rs.Close
    cn.Close
End If

Set rs = Nothing
Set cn = Nothing

' Xu ly du lieu de dua vao PXK theo cong thuc

lr_VLOOKUP_MH = Ton_trong_kho.Range("A" & Rows.Count).End(xlUp).Row ' VLOOKUP(C11,DMHH,5,0)
lr_SUMIFS_NK = Ton_trong_kho.Range("C" & Rows.Count).End(xlUp).Row ' SUMIFS(GHISO!N:N,GHISO!D:D,"NK",GHISO!J:J,C11)
lr_SUMIFS_XK = Ton_trong_kho.Range("E" & Rows.Count).End(xlUp).Row ' SUMIFS(GHISO!N:N,GHISO!D:D,"XK",GHISO!J:J,C11)
lr_Dieukien_MH = Dieukien_MH.Range("A" & Rows.Count).End(xlUp).Row ' Lay cac Ma hang duoc nhap
k = 0

arr_VLOOKUP_MH = Ton_trong_kho.Range("A2:B" & lr_VLOOKUP_MH) ' VLOOKUP(C11,DMHH,5,0)
arr_SUMIFS_NK = Ton_trong_kho.Range("C2:D" & lr_SUMIFS_NK) ' SUMIFS(GHISO!N:N,GHISO!D:D,"NK",GHISO!J:J,C11)
arr_SUMIFS_XK = Ton_trong_kho.Range("E2:F" & lr_SUMIFS_XK) ' SUMIFS(GHISO!N:N,GHISO!D:D,"XK",GHISO!J:J,C11)
ReDim arr_D(1 To lr_VLOOKUP_MH, 1 To 2)
arr_theo_MH = Dieukien_MH.Range("a2:a" & lr_Dieukien_MH)
ReDim nap_arr_theo_MH(1 To lr_Dieukien_MH, 1 To 1)
Set Dic = CreateObject("scripting.dictionary")

For i = LBound(arr_VLOOKUP_MH, 1) To UBound(arr_VLOOKUP_MH, 1)
    If Not Dic.exists(arr_VLOOKUP_MH(i, 1)) Then
        k = k + 1
        Dic.Add arr_VLOOKUP_MH(i, 1), k
        arr_D(i, 1) = arr_VLOOKUP_MH(i, 1)
        arr_D(i, 2) = arr_VLOOKUP_MH(i, 2)
    End If
Next

If lr_GHISO <> 5 Then
    ' Tong so luong da nhap theo Ma hang
    For i = LBound(arr_SUMIFS_NK, 1) To UBound(arr_SUMIFS_NK, 1)
        If Dic.exists(arr_SUMIFS_NK(i, 1)) Then
            j = Dic.Item(arr_SUMIFS_NK(i, 1))
            arr_D(j, 2) = arr_D(j, 2) + arr_SUMIFS_NK(i, 2)
        End If
    Next
    
    ' Tong so luong da xuat theo Ma hang
    For i = LBound(arr_SUMIFS_XK, 1) To UBound(arr_SUMIFS_XK, 1)
        If Dic.exists(arr_SUMIFS_XK(i, 1)) Then
            j = Dic.Item(arr_SUMIFS_XK(i, 1))
            arr_D(j, 2) = arr_D(j, 2) - arr_SUMIFS_XK(i, 2)
        End If
    Next
End If
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
On Error Resume Next
For i = LBound(arr_theo_MH, 1) To UBound(arr_theo_MH, 1)
    If Dic.exists(arr_theo_MH(i, 1)) Then
        j = Dic.Item(arr_theo_MH(i, 1))
        nap_arr_theo_MH(i, 1) = arr_D(j, 2)
    End If
Next

With PXK
    .Unprotect
    .Range("f11").Resize(UBound(arr_theo_MH, 1), 1) = nap_arr_theo_MH
    .Protect
End With
On Error GoTo 0

' Giai phong bo nho
Set arr_VLOOKUP_MH = Nothing
Set arr_SUMIFS_NK = Nothing
Set arr_SUMIFS_XK = Nothing
Set arr_D = Nothing
Set arr_theo_MH = Nothing
Set nap_arr_theo_MH = Nothing
Set Dic = Nothing

Call Tang_toc_code.bat_che_do
End Sub
