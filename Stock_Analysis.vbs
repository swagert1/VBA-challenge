Sub stocks()

Dim k As Integer

k = 1

For k = 1 To Worksheets.Count 'This will iterate through the worksheets

    Worksheets(k).Select

    Dim LastCell As String
    Dim tickers As Variant
    Dim myRange As String
    Dim utickers As Variant
    Dim rows As Integer
    
    
    'Finding the last cell in the ticker column
    
    Range("A2").Select
    Selection.End(xlDown).Select
    LastCell = ActiveCell.Address(False, False)
    
    myRange = "A2:" & LastCell
    
    tickers = Range(myRange)
    
    'Find all unique ticker symbols
    
    utickers = Application.WorksheetFunction.Unique(tickers)
       
    rows = UBound(utickers) - LBound(utickers) + 1
    
    Dim open_date As Double
    Dim close_date As Double
    Dim current_date As Double
    Dim open_price As Double
    Dim close_price As Double
    Dim total_volume As Double
    Dim i As Integer
    
    Dim current_row As Variant
    
    i = 1
    
    current_row = 2
    
    'Iterating through all of the rows and finding the open and close prices for the year, and the total volume
    
    For i = 1 To rows
    
        Cells(i + 1, 9).Value = utickers(i, 1)
        open_date = Cells(200, 2).Value
        
        Do While Cells(current_row, 1).Value = utickers(i, 1)
            
            If Cells(current_row, 1).Value = utickers(i, 1) Then
            
                current_date = Cells(current_row, 2).Value
            
                    If current_date <= open_date Then
                    
                        open_date = current_date
                        open_price = Cells(current_row, 3).Value
                    
                    End If
                
                    If current_date >= close_date Then
                
                        close_date = current_date
                        close_price = Cells(current_row, 6).Value
                    
                    End If
                
                total_volume = total_volume + Cells(current_row, 7).Value
                
            End If
            
            current_row = current_row + 1
            
        Loop
        
        Cells(i + 1, 10).Value = close_price - open_price
        
        'Color code the yearly changes
        
        If Cells(i + 1, 10) > 0 Then
            Cells(i + 1, 10).Interior.ColorIndex = 4
        Else
            Cells(i + 1, 10).Interior.ColorIndex = 3
        End If
        
        Cells(i + 1, 11).Value = (close_price - open_price) / open_price
        
        Cells(i + 1, 12).Value = total_volume
        
        total_volume = 0
        
    Next i
    
    Dim perfromance As Double
    Dim largest_increase As Double
    Dim best_stock As String
    
    Dim largest_decrease As Double
    Dim worst_stock As String
    
    Dim volume As Double
    Dim largest_volume
    Dim most_traded As String
    
    largest_volume = 0
    largest_increase = 0
    largest_decrease = 0
    
    'Finding the best, worst, and total stock volume
    
    For j = 1 To rows
    
        performance = Cells(j + 1, 11)
        volume = Cells(j + 1, 12)
        If performance > largest_increase Then
            largest_increase = performance
            best_stock = Cells(j + 1, 9).Value
        End If
        
        If performance < largest_decrease Then
            largest_decrease = performance
            worst_stock = Cells(j + 1, 9).Value
        End If
        
        If volume > largest_volume Then
            largest_volume = volume
            most_traded = Cells(j + 1, 9).Value
        End If
    
    Next j
    
    Cells(2, 16).Value = best_stock
    Cells(2, 17).Value = largest_increase
    
    Cells(3, 16).Value = worst_stock
    Cells(3, 17).Value = largest_decrease
    
    Cells(4, 16).Value = most_traded
    Cells(4, 17).Value = largest_volume
    
    'Create column/cell labels and tidy up
   
    Cells(1, 9).Value = "Ticker"
    Cells(1, 10).Value = "Yearly Change"
    Cells(1, 11).Value = "Percent Change"
    Cells(1, 12).Value = "Total Stcok Volume"
    Cells(1, 16).Value = "Ticker"
    Cells(1, 17).Value = "Value"
    Cells(2, 15).Value = "Greatest % Increase"
    Cells(3, 15).Value = "Greatest % Decrease"
    Cells(4, 15).Value = "Greatest Total Volume"
    
    Columns("K").NumberFormat = "0.00%"
    Cells(2, 17).NumberFormat = "0.00%"
    Cells(3, 17).NumberFormat = "0.00%"
    
    Columns("I:L").EntireColumn.AutoFit
    Columns("O:Q").EntireColumn.AutoFit
   
Next k
   
End Sub