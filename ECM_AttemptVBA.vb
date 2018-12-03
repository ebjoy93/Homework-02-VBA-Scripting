Sub AttemptVBA()

Dim ticker As String 'holds ticker value
Dim vol As Double 'holds volume total
vol = 0 'initializes vol variable

Dim sumTableRow As Integer
Dim openYear As Double 'holds open year value
Dim yearClose As Double 'holds close year value
sumTableRow = 2 'initializes the variable


For Each ws In Worksheets 'supposed to loop through each worksheet
    
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row 'counts the last row

    Cells(1, 9).Value = "Ticker" 'creates column header
    Cells(1, 10).Value = "Yearly Change" 'creates column header
    Cells(1, 12).Value = "Total Stock Volume" 'creates column header
    Cells(1, 11).Value = "Yearly Percent" 'creates column header

    For x = 2 To 70928 'loops through all the rows in the open year column

        If openYear = 0 Then 'checks values in open year columns
    
            openYear = Cells(x, 3).Value
        
        End If
    
     If Cells(x - 1, 1) = Cells(x, 1) And Cells(x + 1, 1).Value <> Cells(x, 1).Value Then 'checks if cell values are not equal
    
        closeYear = Cells(x, 6).Value
        yearlyChange = closeYear - openYear 'calculates the yearly change between the open and close year
        
        Range("J" & sumTableRow).Value = yearlyChange 'places data in J column
        Range("I" & sumTableRow).Value = ticker 'places data in I column
        Range("K" & sumTableRow).Value = yearlyPercent 'places data in K column
        Range("L" & sumTableRow).Value = vol 'places data in L column
        
        sumTableRow = sumTableRow + 1 'adds another row
        
        vol = 0 'resets vol variable
        
    Else
    
        vol = vol + Cells(x, 7).Value
        
    End If

LastColumn = ws.Cells(1, Columns.Count).End(xlToLeft).Column 'counts the last column
    
    Next x
    
Next ws 'suppoed to move to next worksheet

End Sub
