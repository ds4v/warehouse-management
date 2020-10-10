
Dim Dic As Object
Dim EnableEvents As Boolean
Dim solan As Long

Private Sub CheckBox1_Click()
Dim i As Long
Application.ScreenUpdating = False

With Me.ListBox1
    If .ListCount - 1 <= 3000 Then
    
         If CheckBox1 Then
             For i = 0 To .ListCount - 1
                 .ListIndex = i
                 .Selected(i) = True
             Next
         Else
             For i = 0 To .ListCount - 1
                 EnableEvents = False
                    Dic.RemoveAll
                    .Selected(i) = False
                 EnableEvents = True
             Next
         End If
         
    Else
    
        .ListStyle = fmListStylePlain
        If CheckBox1 Then
            .BackColor = &H8000000D
            '.ForeColor = &H80000005
            For i = 0 To .ListCount - 1
                If Not Dic.exists(.List(i, 0)) Then
                    Dic.Add .List(i, 0), .List(i, 0) & ";" & .List(i, 1) & ";" & _
                    .List(i, 2) & ";" & CDbl(.List(i, 3)) & ";" & CDbl(.List(i, 4))
                End If
            Next
        Else
            .BackColor = &H80000005
            '.ForeColor = &H80000008
            .ListStyle = fmListStyleOption
            Dic.RemoveAll
        End If
        
    End If
End With
Label7 = "Có " & Dic.Count & " mã hàng " & ChrW(273) & ChrW(432) & ChrW(7907) & "c ch" & ChrW(7885) & "n"
Application.ScreenUpdating = True
End Sub
Private Sub CommandButton1_Click()
Dim lr_DMHH As Long, lr As Long, cot As Byte
Dim arr_DMHH As Variant, arr_Dcchon As Variant, arr As Variant
lr_DMHH = DMHH.Range("c" & Rows.Count).End(xlUp).Row
lr = ActiveSheet.Range("c" & Rows.Count).End(xlUp).Row + 1
ReDim arr_Dcchon(0 To lr_DMHH, 1 To 6)
arr_DMHH = DMHH.Range("c6:e" & lr_DMHH)
arr = Dic.items
cot = 4

If Dic.Count = 0 Then
    Application.Assistant.DoAlert "Thông Báo", "B" & ChrW(7841) & "n ch" & ChrW(432) & "a ch" & ChrW(7885) & "n mã hàng nào", 0, 4, 0, 0, 0
    Exit Sub
End If

Call Tang_toc_code.tat_che_do

If ActiveSheet.CodeName = "PNK" Or ActiveSheet.CodeName = "PXK" Then
    Application.EnableEvents = False
    ActiveSheet.Unprotect
    If Dic.Count > 0 Then
        For i = 0 To Dic.Count - 1
            arr_Dcchon(i, 1) = Left(arr(i), InStr(1, arr(i), ";", vbTextCompare) - 1)
            arr_Dcchon(i, 2) = Mid(arr(i), Len(arr_Dcchon(i, 1)) + 2)
            arr_Dcchon(i, 2) = Left(arr_Dcchon(i, 2), InStr(1, arr_Dcchon(i, 2), ";", vbTextCompare) - 1)
            arr_Dcchon(i, 3) = Mid(arr(i), Len(arr_Dcchon(i, 1)) + 2, Len(arr(i)) - Len(arr_Dcchon(i, 1)))
            arr_Dcchon(i, 3) = Mid(arr_Dcchon(i, 3), Len(arr_Dcchon(i, 2)) + 2)
            arr_Dcchon(i, 3) = Left(arr_Dcchon(i, 3), InStr(1, arr_Dcchon(i, 3), ";", vbTextCompare) - 1)
            arr_Dcchon(i, 4) = Right(arr(i), Len(arr(i)) - Len(arr_Dcchon(i, 1)) - Len(arr_Dcchon(i, 2)) - Len(arr_Dcchon(i, 3)) - 3)

            If ActiveSheet.CodeName = "PXK" Then cot = 6
            
            arr_Dcchon(i, cot) = Left(arr_Dcchon(i, 4), InStr(1, arr_Dcchon(i, 4), ";", vbTextCompare) - 1)
            arr_Dcchon(i, 5) = Right(arr(i), Len(arr(i)) - Len(arr_Dcchon(i, 1)) - Len(arr_Dcchon(i, 2)) - Len(arr_Dcchon(i, 3)) - Len(arr_Dcchon(i, cot)) - 4)
        Next i
        
        If ActiveSheet.CodeName = "PNK" Then cot = 5
        ActiveSheet.Range("c" & lr).Resize(UBound(arr_Dcchon, 1) + 1, cot) = arr_Dcchon
        
    End If
    
    If ActiveSheet.CodeName = "PXK" Then
        Call Ham_chung.copy_sheet(PXK, Dieukien_MH, 11, "c", 1)
        Call Ton_kho_PXK.kho_co_de_xuat
    End If
    
    Set Dic = Nothing
    Erase arr
    Unload Me
    ActiveSheet.Protect
    Application.EnableEvents = True
End If

Call Tang_toc_code.bat_che_do
End Sub

Private Sub ListBox1_change()
Dim id

If EnableEvents Then
    With Me.ListBox1
        id = .ListIndex
        
        If .ListCount - 1 > 3000 Then
            If .Selected(id) And CheckBox1 Then
            
                Dic.Remove (.List(id, 0))
                Exit Sub
                
            ElseIf .Selected(id) = False And CheckBox1 Then
                solan = solan + 1
                If solan = 1 Then
                    Application.Assistant.DoAlert "Thông Báo", "D" & ChrW(7919) & " li" & ChrW(7879) & "u " & _
                    ChrW(273) & ChrW(432) & ChrW(7907) & "c ch" & ChrW(7885) & "n l" & ChrW( _
                    7847) & "n th" & ChrW(7913) & " 2 s" & ChrW(7869) & " n" & ChrW(7857) & _
                    "m " & ChrW(7903) & " dòng cu" & ChrW(7889) & "i cùng", 0, 4, 0, 0, 0
                End If
                
                If Not Dic.exists(.List(id, 0)) Then
                    Dic.Add .List(id, 0), .List(id, 0) & ";" & .List(id, 1) & ";" & _
                    .List(id, 2) & ";" & CDbl(.List(id, 3)) & ";" & CDbl(.List(id, 4))
                End If
                Exit Sub
            End If
        End If
        
        If .Selected(id) Then
            If Not Dic.exists(.List(id, 0)) Then
                Dic.Add .List(id, 0), .List(id, 0) & ";" & .List(id, 1) & ";" & _
                .List(id, 2) & ";" & CDbl(.List(id, 3)) & ";" & CDbl(.List(id, 4))
            End If
        Else
            On Error Resume Next
            Dic.Remove (.List(id, 0))
            On Error GoTo 0
        End If
    End With
End If
Label7 = "Có " & Dic.Count & " mã hàng " & ChrW(273) & ChrW(432) & ChrW(7907) & "c ch" & ChrW(7885) & "n"
End Sub
Private Sub TextBox1_change()
strsearch = LCase(TextBox1.Text)
CheckBox1 = False
Dim lr As Long
lr = DMHH.Range("c" & Rows.Count).End(xlUp).Row

' Neu du lieu <= 3000 thi listbox se loc ra nhung thang thoa dieu kien,
' cach nay de tien quan sat va thao tac cho nhanh nhung neu du lieu lon lam cach nay se chay rat lau,
' vi vay khi du lieu > 3000 se thuc hien trich loc bang SQL

If lr - 5 <= 3000 Then
    Dim rw As Range, r As Range
    Set r = DMHH.Range("c6:h" & lr)
    With ListBox1
        .Clear ' Lam moi Listbox
        For Each rw In r.Rows
            ' Do tim theo Ma hang, ten hang, don vi tinh
            If InStr(LCase(DMHH.Cells(rw.Row, 3) & DMHH.Cells(rw.Row, 4) & DMHH.Cells(rw.Row, 5)), strsearch) Then
                ' Neu chuoi can tim co trong chuoi Ma hang nôi voi Ten hang nôi voi Don vi tinh thi add vao Listbox
                .AddItem DMHH.Cells(rw.Row, 1).Value
                .List(ListBox1.ListCount - 1, 0) = DMHH.Cells(rw.Row, 3).Value
                .List(ListBox1.ListCount - 1, 1) = DMHH.Cells(rw.Row, 4).Value
                .List(ListBox1.ListCount - 1, 2) = DMHH.Cells(rw.Row, 5).Value
                .List(ListBox1.ListCount - 1, 3) = DMHH.Cells(rw.Row, 6).Value
                .List(ListBox1.ListCount - 1, 4) = DMHH.Cells(rw.Row, 7).Value
                .List(ListBox1.ListCount - 1, 5) = DMHH.Cells(rw.Row, 8).Value
            End If
        Next rw
    End With
End If

End Sub

Private Sub CommandButton2_Click()
Call Tang_toc_code.tat_che_do

Dim cn As Object, rs As Object
Dim sql As String

Set rs = Nothing
Set cn = Nothing
Set cn = CreateObject("adodb.connection")
Set rs = CreateObject("adodb.recordset")
strsearch = LCase(TextBox1.Text)

cn.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & _
ThisWorkbook.FullName & ";Extended Properties=""Excel 12.0 Xml;HDR=YES;IMEX=1"";"
cn.Open

' Truy van theo Ten cua ma hang, phai dung sheet moi chu neu dung name dong hoac 1 vung de truy van thi chi thao tac duoc voi 65356 dong
sql = "SELECT * FROM [copy_DMHH$] WHERE ten_hang LIKE ""%" & strsearch & "%"""
rs.Open sql, cn

If Not (rs.BOF And rs.EOF) Then
    ListBox1.List = Ham_chung.TransposeArray(rs.GetRows())
Else
    ListBox1.Clear
End If
Call Tang_toc_code.bat_che_do
End Sub

Private Sub UserForm_Initialize()
Dim lr As Long, solan As Long, sophieu As String
lr = DMHH.Range("c" & Rows.Count).End(xlUp).Row
Set Dic = CreateObject("Scripting.Dictionary")
ListBox1.List = DMHH.Range("c6:h" & lr).Value ' Nap toan bo du lieu cho list box
CommandButton2.Visible = False ' An nut tim kiem vi chua biet du lieu co > 3000 dong ko
EnableEvents = True
sophieu = PNK.Range("i2")

' Neu du lieu hon 3000 dong thi se xuat hien nut tim kiem de tim kiem bang SQL chu neu de listbox tim kiem thi se rat cham
If ListBox1.ListCount > 3000 Then CommandButton2.Visible = True
If ActiveSheet.CodeName = "PXK" Then sophieu = PXK.Range("j2")
Label7 = "Có " & Dic.Count & " mã hàng " & ChrW(273) & ChrW(432) & ChrW(7907) & "c ch" & ChrW(7885) & "n" ' Hien so Ma hang duoc chon
Label8 = "S" & ChrW(7889) & " phi" & ChrW(7871) & "u: " & sophieu ' Hien thi so phieu hien tai
solan = 1
End Sub
