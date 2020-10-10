Sub THNXT_Capnhat()
Call Tang_toc_code.tat_che_do
THNXT.Unprotect
THNXT.Range("B10:B" & Rows.Count).EntireRow.Hidden = False
''''''''''''''''''''''''''''''''''''''' Sô luong tôn dâu '''''''''''''''''''''''''''''''''''''''''''''
Dim lr_DMHH As Long, lr_GHISO As Long
Dim Tu_ngay As Date, Den_ngay As Date
Dim i As Long, k As Long, j As Long
Dim arr_DMHH, arr_GHISO, arr_Tungay, arr_D, arr_Tondau
Dim Dic As Object

k = 0
lr_DMHH = DMHH.Range("C" & Rows.Count).End(xlUp).Row
lr_GHISO = GHISO.Range("J" & Rows.Count).End(xlUp).Row
Tu_ngay = THNXT.Range("i5").Value
Den_ngay = THNXT.Range("k5").Value

arr_DMHH = DMHH.Range("c6:c" & lr_DMHH)
arr_GHISO = GHISO.Range("c6:p" & lr_GHISO)

ReDim arr_Tungay(1 To lr_GHISO, 1 To 3)
ReDim arr_D(1 To lr_GHISO, 1 To 2)
ReDim arr_Tondau(1 To UBound(arr_DMHH, 1), 1 To 2)
Set Dic = CreateObject("scripting.dictionary")

' Loc ra nhung mat hang < Tu_ngay
For i = 1 To UBound(arr_GHISO, 1)
    If arr_GHISO(i, 4) < Tu_ngay Then
        k = k + 1
        arr_Tungay(k, 1) = arr_GHISO(i, 2)
        arr_Tungay(k, 2) = arr_GHISO(i, 8)
        arr_Tungay(k, 3) = arr_GHISO(i, 12)
    End If
Next

k = 0
For i = 1 To UBound(arr_GHISO, 1)
    If Not (arr_Tungay(i, 1) = Empty And arr_Tungay(i, 2) = Empty) Then
        If Not Dic.exists(arr_Tungay(i, 1) & "-" & arr_Tungay(i, 2)) Then
            k = k + 1
            Dic.Add arr_Tungay(i, 1) & "-" & arr_Tungay(i, 2), k
            arr_D(k, 1) = arr_Tungay(i, 2)
            arr_D(k, 2) = arr_Tungay(i, 3)
        Else
            j = Dic.Item(arr_Tungay(i, 1) & "-" & arr_Tungay(i, 2))
            arr_D(j, 2) = arr_D(j, 2) + arr_Tungay(i, 3)
        End If
    Else
        Exit For
    End If
Next

' Loc cac ma hang xuât kho, nhâp kho va dua vao arr_Ton_dau
For i = 1 To UBound(arr_DMHH, 1)
    If Dic.exists("NK" & "-" & arr_DMHH(i, 1)) Then
        j = Dic.Item("NK" & "-" & arr_DMHH(i, 1))
        arr_Tondau(i, 1) = arr_D(j, 2)
    End If
    
    If Dic.exists("XK" & "-" & arr_DMHH(i, 1)) Then
        j = Dic.Item("XK" & "-" & arr_DMHH(i, 1))
        arr_Tondau(i, 2) = arr_D(j, 2)
    End If
Next
THNXT.Range("G10:H" & Rows.Count).ClearContents
THNXT.Range("G10").Resize(UBound(arr_DMHH, 1), 2) = arr_Tondau

'''''''''''''''''''''''''''''''''' Sô luong nhap xuat kho ''''''''''''''''''''''''''''''''''''''''''''''''''
Dim cn As Object, rs As Object
Dim Dulieu As String

lr_GHISO = GHISO.Range("D" & Rows.Count).End(xlUp).Row
Tu_ngay = Format(THNXT.Range("I5"), "mm/dd/yyyy") ' Format theo dinh dang trong SQL
Den_ngay = Format(THNXT.Range("K5"), "mm/dd/yyyy")

Set cn = CreateObject("ADODB.Connection")
Set rs = CreateObject("ADODB.Recordset")

cn.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & Chr(39) & _
ThisWorkbook.FullName & Chr(39) & ";Extended Properties=""Excel 12.0 Xml;HDR=YES;IMEX=1"";"
cn.Open

' Xoa du lieu cu di
Xuly_NXT.Range("A2:D" & Rows.Count).ClearContents
THNXT.Range("J10:J150009").ClearContents ' Xoa du lieu cu di
THNXT.Range("L10:L150009").ClearContents

If lr_GHISO <> 5 Then
    ' Xu Ly du lieu nhap kho theo công thuc SUMIFS(GHISO!N:N,GHISO!F:F,">="&$I$5,GHISO!F:F,"<="&$K$5,GHISO!J:J,$C10,GHISO!D:D,"NK")
    Set rs = cn.Execute("SELECT ma_hang, SUM(so_luong) FROM [copy_GHISO$] WHERE loai ='NK' AND ngay_lap >= #" & _
                        Tu_ngay & "# AND ngay_lap <= #" & Den_ngay & "# GROUP BY ma_hang")
    Xuly_NXT.Range("A2").CopyFromRecordset rs
    
    ' Dua du lieu nhap kho vao sheet THNXT
    Set rs = cn.Execute("SELECT IIF(ISNULL(SL_nhap), 0, SL_nhap * 1) FROM [Xuly_NXT$] A RIGHT JOIN [copy_DMHH$] B on A.MH_nhap = B.ma_hang ORDER BY B.STT")
    THNXT.Range("J10").CopyFromRecordset rs
    
    ' Xu Ly du lieu xuat kho theo công thuc SUMIFS(GHISO!N:N,GHISO!F:F,">="&$I$5,GHISO!F:F,"<="&$K$5,GHISO!J:J,$C10,GHISO!D:D,"XK")
    Set rs = cn.Execute("SELECT ma_hang, SUM(so_luong) FROM [copy_GHISO$] WHERE loai ='XK' AND ngay_lap >= #" & _
                        Tu_ngay & "# AND ngay_lap <= #" & Den_ngay & "# GROUP BY ma_hang")
    Xuly_NXT.Range("C2").CopyFromRecordset rs
    
    ' Dua du lieu xuat kho vao sheet THNXT
    Set rs = cn.Execute("SELECT IIF(ISNULL(SL_xuat), 0, SL_xuat * 1) FROM [Xuly_NXT$] A RIGHT JOIN [copy_DMHH$] B on A.MH_xuat = B.ma_hang ORDER BY B.STT")
    THNXT.Range("L10").CopyFromRecordset rs
    
    rs.Close
    cn.Close
End If

Set rs = Nothing
Set cn = Nothing
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

' An nhung dong ko co du lieu
If lr_DMHH + 5 <= 150009 Then THNXT.Range("B" & lr_DMHH + 5 & ":B150009").EntireRow.Hidden = True
THNXT.Protect
Call Tang_toc_code.bat_che_do
End Sub

Sub THNXT_Xuat_bao_cao()
THNXT.PrintOut
End Sub
