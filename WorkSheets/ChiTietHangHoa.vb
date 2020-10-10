Private Sub Worksheet_Change(ByVal Target As Range)
Call Tang_toc_code.tat_che_do
Dim cn As Object, rs As Object
Dim DK_loc As String, lr_GHISO As Long
Dim Tu_ngay As Date, Den_ngay As Date

CTHH.Unprotect
DK_loc = CTHH.Range("F6").Value
lr_GHISO = GHISO.Range("D" & Rows.Count).End(xlUp).Row

Tu_ngay = Format(CTHH.Range("E5").Value, "mm/dd/yyyy") ' Format theo dinh dang trong SQL
Den_ngay = Format(CTHH.Range("G5").Value, "mm/dd/yyyy")

If Not Intersect([F6], Target) Is Nothing Then
    CTHH.Range("B13:I150012").ClearContents
    CTHH.Range("B13:B" & Rows.Count).EntireRow.Hidden = False
    
    Set cn = CreateObject("ADODB.Connection")
    Set rs = CreateObject("ADODB.Recordset")
    
    cn.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & Chr(39) & _
    ThisWorkbook.FullName & Chr(39) & ";Extended Properties=""Excel 12.0 Xml;HDR=YES;IMEX=1"";"
    cn.Open

    Set rs = cn.Execute("SELECT so_phieu, FORMAT(ngay_lap, 'dd-mm-yyyy'), dien_giai, so_luong, don_gia, thanh_tien, ghi_chu FROM [copy_GHISO$] WHERE ma_hang = '" & _
                        DK_loc & "' AND ngay_lap >= #" & Tu_ngay & "# AND ngay_lap <= #" & Den_ngay & "#")
    CTHH.Range("C13").CopyFromRecordset rs
    
    rs.Close
    cn.Close
    
    '''''''' Dien so thu tu de sap xep du lieu khi dua vao sheet THNXT '''''''
    Dim arr_STT, i As Long, lr_copy_DMHH As Long
    lr_CTHH = CTHH.Range("D" & Rows.Count).End(xlUp).Row - 12
    If lr_CTHH > 0 Then
        ReDim arr_STT(1 To lr_CTHH, 1 To 1)
        For i = 1 To lr_CTHH
            arr_STT(i, 1) = i
        Next
        CTHH.Range("B13").Resize(UBound(arr_STT), 1) = arr_STT
    End If
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    If lr_CTHH <= 150000 Then CTHH.Range("B" & lr_CTHH + 13 & ":B150012").EntireRow.Hidden = True
End If

Set rs = Nothing
Set cn = Nothing

''''''''''''''''''' Ton dau, Nhap kho, Xuat kho, Ton cuoi '''''''''''''''''''''''''''
Dim arr_THNXT
Dim j As Long, lr_DMHH As Long
lr_DMHH = DMHH.Range("C" & Rows.Count).End(xlUp).Row

ReDim arr_THNXT(1 To lr_DMHH - 5, 1 To 13)
arr_THNXT = THNXT.Range("C10:O" & lr_DMHH)
For i = LBound(arr_THNXT, 1) To UBound(arr_THNXT, 1)
    If DK_loc = arr_THNXT(i, 1) Then
        With CTHH
            .Range("C5") = arr_THNXT(i, 4)
            .Range("D5") = arr_THNXT(i, 7)
            .Range("C6") = arr_THNXT(i, 8)
            .Range("D6") = arr_THNXT(i, 9)
            .Range("C7") = arr_THNXT(i, 10)
            .Range("D7") = arr_THNXT(i, 11)
            .Range("C8") = arr_THNXT(i, 12)
            .Range("D8") = arr_THNXT(i, 13)
        End With
        Exit For
    End If
Next
Set arr_THNXT = Nothing
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
CTHH.Protect
Call Tang_toc_code.bat_che_do
End Sub

Sub CTHH_Xuat_bao_cao()
Call Tang_toc_code.tat_che_do
Application.EnableEvents = True

Dim arr_DMHH, lr_DMHH As Long
Dim PrintFrom As String, PrintTo As String
Dim i As Long, Index_From As Long, Index_To As Long
lr_DMHH = DMHH.Range("c" & Rows.Count).End(xlUp).Row
ReDim arr_DMHH(1 To lr_DMHH, 1 To 2)
arr_DMHH = DMHH.Range("B6:C" & lr_DMHH)

PrintFrom = CTHH.Range("K5")
PrintTo = CTHH.Range("K7")

For i = LBound(arr_DMHH, 1) To UBound(arr_DMHH, 1)
    If PrintFrom = arr_DMHH(i, 2) Then
        Index_From = arr_DMHH(i, 1)
        Exit For
    End If
Next

For i = LBound(arr_DMHH, 1) To UBound(arr_DMHH, 1)
    If PrintTo = arr_DMHH(i, 2) Then
        Index_To = arr_DMHH(i, 1)
        Exit For
    End If
Next

If Index_From <= Index_To Then
    For i = Index_From To Index_To
        CTHH.Range("F6") = arr_DMHH(i, 2)
        CTHH.PrintOut
    Next
Else
    Application.Assistant.DoAlert "C" & ChrW(7843) & "nh Báo", "Th" & ChrW(432) & ChrW(769) & " t" & ChrW(432) & _
    ChrW(803) & " ba" & ChrW(803) & "n cho" & ChrW(803) & "n không h" & ChrW _
    (417) & ChrW(803) & "p lê" & ChrW(803), 0, 1, 0, 0, 0
End If

Call Tang_toc_code.bat_che_do
End Sub
