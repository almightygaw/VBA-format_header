Attribute VB_Name = "format_header"
Sub format_header()
Attribute format_header.VB_ProcData.VB_Invoke_Func = "w\n14"

    Application.ScreenUpdating = False

' autofit columns, set font, font size
    Dim i As Range
    Dim maxWidth As Integer
    maxWidth = 50  ' set maxWidth = 50

    For Each i In ActiveSheet.UsedRange.Rows(1).Cells
        i.EntireColumn.Font.Name = "Calibri"
        i.EntireColumn.Font.Size = 11
        i.EntireColumn.AutoFit
        If i.EntireColumn.ColumnWidth > maxWidth _
            Then i.EntireColumn.ColumnWidth = maxWidth
    Next i
    
' autofit row height
    Dim j As Range
    Dim maxHeight As Integer
    maxHeight = 15 ' set maxHeight = 15
    
    For Each j In ActiveSheet.UsedRange.Cells
    j.EntireRow.AutoFit
        If j.EntireRow.RowHeight > maxHeight _
            Then j.EntireRow.RowHeight = maxHeight
    Next j
        
' highlight, bold, center header row
    Dim k As Range
    
    For Each k In ActiveSheet.UsedRange.Rows(1).Cells
        k.Font.Bold = True
        k.Interior.ColorIndex = 36
        k.HorizontalAlignment = xlCenter
        k.BorderAround ColorIndex:=1, Weight:=xlThin
    Next k
    
' freeze header row
    Rows("2:2").Select
    ActiveWindow.FreezePanes = True
    
    Range("A1").Select
    
    Application.ScreenUpdating = True
    
End Sub



