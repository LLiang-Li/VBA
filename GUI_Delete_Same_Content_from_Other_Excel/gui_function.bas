Attribute VB_Name = "gui_function"
Sub gui_delete_row(file_name As String, row_now As String, row As String, t As Boolean)

    'Dim file_name As String
    
    
    'file_name = "小表.xlsx"
    
    Dim now_file_name As String
    
    now_file_name = ActiveWorkbook.Name
    
    'MsgBox now_file_name
    
    
    Windows(file_name).Activate
    
    
    
    '获取小表的行数
    Dim file_name_row As Long
    file_name_row = Range(row & "65535").End(xlUp).row - 1
    
    Dim file_num(65536) As String
    
    Dim num As Long
    
    
    For num = 1 To file_name_row
        file_num(num) = Range(row & num + 1)
    Next
    
    '获取大表的行数
    Windows(now_file_name).Activate
    
    Dim now_file_name_row As Long
    now_file_name_row = Range(row_now & "65535").End(xlUp).row - 1
    
    Dim now_file_num(65536) As String
    
    Dim now_num As Long
    
    
    For now_num = 1 To now_file_name_row
        now_file_num(now_num) = Range(row_now & now_num + 1)
    Next
    
    Dim i, j As Long
    
    For i = now_file_name_row To 1 Step -1
        For j = 1 To file_name_row
            
            If now_file_num(i) = file_num(j) Then
                Rows(Str(i + 1)).Select
                Selection.Delete Shift:=xlUp
                Exit For
            End If
        Next
    Next
    
    If t = True Then
        now_file_name_row = Range(row_now & "65535").End(xlUp).row - 1
        
        For now_num = 1 To now_file_name_row
            Range("A" & now_num + 1) = now_num
        Next
    End If
    

End Sub



Sub call_chuanti()

Delete_the_same_col.Show

End Sub













