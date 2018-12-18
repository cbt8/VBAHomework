Attribute VB_Name = "Module1"
Sub Stocks()


Dim sum As Double
Dim Count As Long
Dim n As Long

Count = 2
sum = 0
n = 70926

'''n = Worksheets("A").Range("A:A").Cells.SpecialCells(xlCellTypeConstants).Count'''

For i = 2 To n

    sum = sum + Cells(i, 7).Value
                    
        If Cells(i, 1).Value <> Cells(i + 1, 1).Value Then
            
            Cells(Count, 11).Value = Cells(i, 1).Value
            Cells(Count, 12).Value = sum
            sum = 0
            Count = Count + 1
        
        End If
           
Next
     
End Sub
