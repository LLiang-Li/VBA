VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Delete_the_same_col 
   Caption         =   "删除重复信息"
   ClientHeight    =   2260
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   7390
   OleObjectBlob   =   "Delete_the_same_col.frx":0000
   StartUpPosition =   1  '所有者中心
End
Attribute VB_Name = "Delete_the_same_col"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()

Dim file_name As String

file_name = Delete_the_same_col.file_name

Dim t As Boolean

t = Delete_the_same_col.file_col

Dim row_now As String

row_now = Delete_the_same_col.now_row

Dim row As String

row = Delete_the_same_col.file_row

If Right(file_name, 3) = "xls" Or Right(file_name, 4) = "xlsx" Then
    
    Call gui_delete_row(file_name, row_now, row, t)
    
Else
    MsgBox "文件名不是EXCEL格式！！"
End If

Unload Me

End Sub

'''''''''''''''
