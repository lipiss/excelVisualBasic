Sub showAllNumbersOfSelectedArea()
''https://www.ablebits.com/office-addins-blog/custom-excel-number-format/
'for large numbers that display with scientific notation set format to just '#'
'Worksheets("Sheet2").Activate
'ActiveSheet.Columns(1).Select 'or Worksheets("SheetName").Range("A:A").Select
'hash mark
Selection.NumberFormat = "#"
End Sub


Sub vba_borders()
Dim myRange As Range
Set myRange = Selection
Selection.Address _
        .Borders(xlEdgeBottom) _
            .LineStyle = XlLineStyle.xlContinuous

End Sub

Sub SelectRange()
Range("A1:D20").Select
End Sub
