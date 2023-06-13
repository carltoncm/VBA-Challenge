Sub multiple_year()

    'Declaring variables
    Dim i As Long
    Dim ws As Worksheet
    Dim summaryrow As Integer
    Dim lastrow As Long
    
    'Loop through all worksheets
    For Each ws In ThisWorkbook.Worksheets
    
        'Create summary table header and set row values
        ws.Range("K1").Value = "Ticker"
        ws.Range("L1").Value = "Yearly Change"
        ws.Range("M1").Value = "Percentage Change"
        ws.Range("N1").Value = "Total Stock Volume"
        summaryrow = 2
        lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
        'Loop through all rows per worksheet
        For i = 2 To lastrow
        
        'Declaring variables
        Dim tickername As String
        Dim open_closed As Double
        Dim yearchange As Double
        Dim tsvolume As Double
        
        tickername = ws.Cells(i, 1).Value
        
            'Setting conditions for the loop
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                
                'Placing each ticker name in the summary table
                ws.Range("K" & summaryrow).Value = tickername
                
                'Totaling the difference between each tickers open and closed values
                open_closed = ws.Cells(i, 3).Value - ws.Cells(i, 6).Value
                yearchange = open_closed + yearchange
                ws.Range("L" & summaryrow).Value = yearchange
                
                
                'Totaling the Total Stock Volume for each ticker name
                tsvolume = ws.Cells(i, 7).Value + tsvolume
                ws.Range("N" & summaryrow).Value = tsvolume
                
                'Resetting the values for the next loop
                summaryrow = summaryrow + 1
                tsvolume = 0
                open_closed = 0
                yearchange = 0
                percentchange = 0
                
            Else
            
                'Adding like cells
                open_closed = ws.Cells(i, 3).Value - ws.Cells(i, 6).Value
                yearchange = open_closed + yearchange
                tsvolume = ws.Cells(i, 7).Value + tsvolume
                
                
            End If
            
        Next i
        
        'Set up loop for percent change column
        For i = 2 To lastrow
        
        Dim opening As Double
        Dim closerate As Double
        Dim percentdiff As Double
        
        summaryrow = 2
        
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                closerate = ws.Cells(i, 6).Value
                opening = ws.Cells(i, 3).Value
                percentdiff = (opening - closerate) / opening
                ws.Range("M" & summaryrow).Value = percentdiff
                ws.Range("M" & summaryrow).NumberFormat = "0.00%"
            
            Else
            
                summaryrow = summaryrow + 1
            
            End If
        
        Next i
        
        'Creating a new lastrow to meet conditions for color formatting on each ws
        lastrow = ws.Range("L" & Rows.Count).End(xlUp).Row
        
        'Set color condition for yearly change column
        For i = 2 To lastrow
        
            If ws.Cells(i, 12).Value > 0 Then
                ws.Cells(i, 12).Interior.ColorIndex = 4
                
            Else
                ws.Cells(i, 12).Interior.ColorIndex = 3
                
            End If
            
        Next i
        
            
    Next ws

End Sub


