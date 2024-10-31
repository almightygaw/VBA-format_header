Attribute VB_Name = "format_header"
Sub format_header()
Attribute format_header.VB_ProcData.VB_Invoke_Func = "w\n14"

    Application.ScreenUpdating = False
    
' clear existing formatting
    With ActiveSheet.UsedRange
      Selection.Interior.Color = xlNone
      Selection.Borders.LineStyle = xlNone
    End With
    
' set font, font size, font style
    With ActiveSheet.UsedRange.Font
      .Name = "Calibri"
      .Color = vbBlack
      .Size = 11
    End With
      
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

' autofit columns, set font, font size
    ActiveSheet.UsedRange.Columns.AutoFit
  
' autofit row height
    ActiveSheet.UsedRange.Rows.AutoFit
    
' border UsedRange
    With ActiveSheet.UsedRange.Borders
      .LineStyle = xlContinuous
      .Color = vbBlack
      .Weight = xlThin
    End With
       
    Range("A1").Select
    
    Application.ScreenUpdating = True
    
End Sub



