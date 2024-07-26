Sub Find_Bond_Matches()
Dim CompareRange As Variant, x As Variant, y As Variant
' Set CompareRange equal to the range to which you will
' compare the selection.
Dim SelectionRange As Variant
Set SelectionRange = Range("E1:E34")
Set CompareRange = Range("I1:I536")
' NOTE: If the compare range is located on another workbook
' or worksheet, use the following syntax.
' Set CompareRange = Workbooks("Book2"). _
' Worksheets("Sheet2").Range("C1:C5")
'
' Loop through each cell in the selection and compare it to
' each cell in CompareRange.
For Each x In SelectionRange
For Each y In CompareRange
If x = y Then x.Offset(0, 1) = y.Offset(0, -2)
Next y
Next x
End Sub