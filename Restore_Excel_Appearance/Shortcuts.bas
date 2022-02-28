Attribute VB_Name = "Shortcuts"
Sub set_all_sheets_100()

Dim sht As Worksheet

For Each sht In ActiveWorkbook.Worksheets
    
     sht.Select
     On Error GoTo ERR_1
     ActiveWindow.Zoom = 100
     Range("A1").Select
ERR_1:
Next

Sheets(1).Select

End Sub
