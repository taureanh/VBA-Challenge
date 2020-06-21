Sub Easy_Solution()
    
    'Declare and set worksheet and define variables
    Dim ws          As Worksheet
    Dim x           As Double
    Dim stock_volume As Double
    
    'Loop through all stocks for one year
    For Each ws In Worksheets
        
        ws.Cells(1, 9).Value = Cells(1, 1).Value
        ws.Cells(1, 10).Value = "Total Stock Value"
        
        x = 2
        
        ws.Cells(x, 9).Value = ws.Cells(x, 1).Value
        
        'Define last row of worksheet
        
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        'Ticker symbol output
        
        For i = 2 To LastRow
            
            If ws.Cells(i, 1).Value = ws.Cells(x, 9).Value Then
                
                stock_volume = stock_volume + ws.Cells(i, 7).Value
                
            Else
                
                ws.Cells(x, 10).Value = stock_volume
                
                stock_volume = ws.Cells(i, 7).Value
                
                'add counter
                
                x = x + 1
                ws.Cells(x, 9).Value = ws.Cells(i, 1).Value
                
            End If
            
        Next i
        
        ws.Cells(x, 10).Value = stock_volume
        
    Next ws
    
End Sub

Sub Moderate_Solution()
    
    Dim open_price  As Variant
    Dim close_price As Variant
    Dim i           As Double
    Dim ws          As Worksheet
    
    Dim x           As Double
    
    For Each ws In Worksheets
        
        'Create column headings
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        
        x = 2
        i = 2
        
        ws.Cells(x, 9).Value = ws.Cells(x, 1).Value
        
        open_price = ws.Cells(i, 3).Value
        
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        For i = 2 To LastRow
            
            'Calculate yearly and percentage change
            
            If ws.Cells(i, 1).Value = ws.Cells(x, 9).Value Then
                
                stock_volume = stock_volume + ws.Cells(i, 7).Value
                
                close_price = ws.Cells(i, 6).Value
                
                'need to account for divide by zero condition causing error
                
            Else
                
                ws.Cells(x, 10).Value = close_price - open_price
                
                If close_price <= 0 Then
                    
                    ws.Cells(x, 11).Value = 0
                    
                Else
                    
                    ws.Cells(x, 11).Value = (close_price / open_price) - 1
                    
                End If
                
                'Format cells with color
                
                ws.Cells(x, 11).Style = "Percent"
                
                If ws.Cells(x, 10).Value >= 0 Then
                    
                    ws.Cells(x, 10).Interior.ColorIndex = 4
                    
                Else
                    
                    ws.Cells(x, 10).Interior.ColorIndex = 3
                    
                End If
                
                ws.Cells(x, 12).Value = stock_volume
                
                open_price = ws.Cells(i, 3).Value
                
                stock_volume = ws.Cells(i, 7).Value
                
                x = x + 1
                ws.Cells(x, 9).Value = ws.Cells(i, 1).Value
                
            End If
            
        Next i
        
        ws.Cells(x, 10).Value = close_price - open_price
        
        If close_price <= 0 Then
            
            ws.Cells(x, 11).Value = 0
            
        Else
            
            ws.Cells(x, 11).Value = (close_price / open_price) - 1
            
        End If
        
        ws.Cells(x, 11).Style = "Percent"
        
        If ws.Cells(x, 10).Value >= 0 Then
            
            ws.Cells(x, 10).Interior.ColorIndex = 4
            
        Else
            
            ws.Cells(x, 10).Interior.ColorIndex = 3
            
        End If
        
        ws.Cells(x, 12).Value = stock_volume
        
    Next ws
    
End Sub

Sub Hard_solution()
    
    Dim ws          As Worksheet
    For Each ws In Worksheets
        
        'Create column headers for summary table
        
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total Volume"
        
        Dim Ticker_Greatest_Increase As Variant
        Dim Volume_Greatest_Increase As Variant
        Dim Ticker_Greatest_Decrease As Variant
        Dim Ticker_Greatest_Total_Volume As Variant
        Dim Volume_Greatest_Total_Volume As Variant
        
        'Create row names and format cells
        ws.Cells(2, 16).Value = Ticker_Greatest_Increase
        ws.Cells(2, 17).Value = Volume_Greatest_Increase
        ws.Cells(2, 17).Style = "Percent"
        ws.Cells(3, 16).Value = Ticker_Greatest_Decrease
        ws.Cells(3, 17).Value = Volume_Greatest_Decrease
        ws.Cells(3, 17).Style = "Percent"
        ws.Cells(4, 16).Value = Ticker_Greatest_Total_Volume
        ws.Cells(4, 17).Value = Volume_Greatest_Total_Volume
        
        Volume_Greatest_Decrease = 100000
        Ticker_Greatest_Decrease = 100000
        
        LastRow = ws.Cells(Rows.Count, 9).End(xlUp).Row
        
        For x = 2 To LastRow
            
            If ws.Cells(x, 11).Value > Volume_Greatest_Increase Then
                
                Ticker_Greatest_Increase = ws.Cells(x, 9).Value
                Volume_Greatest_Increase = ws.Cells(x, 11).Value
                
            End If
            
            If ws.Cells(x, 11).Value < Volume_Greatest_Decrease Then
                
                Ticker_Greatest_Decrease = ws.Cells(x, 9).Value
                Volume_Greatest_Decrease = ws.Cells(x, 11).Value
                
            End If
            
            If ws.Cells(x, 12).Value > Volume_Greatest_Total_Volume Then
                
                Ticker_Greatest_Total_Volume = ws.Cells(x, 9).Value
                Volume_Greatest_Total_Volume = ws.Cells(x, 12).Value
                
            End If
            
        Next x
        
    Next ws
    
End Sub