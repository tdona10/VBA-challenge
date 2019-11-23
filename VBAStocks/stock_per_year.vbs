Sub stock_per_year()
    Dim WS     As Worksheet
    'To do for each sheet
    For Each WS In ThisWorkbook.Worksheets
        WS.Activate
        ' Declare stock related variables
        Dim Ticker_Symbol As String
        Dim Open_Price As Double
        Dim Close_Price As Double
        Dim Yearly_Change As Double
        Dim Percent_Change As Double
        
        'Declare volume of stock variables and initialize it
        Dim Stock_Volume As Double
        Stock_Volume = 0
        
        'Keep track of the location for each stock name in the summary table
        Dim Summary_Table_Row As Long
        Summary_Table_Row = 2
        
        ' Summary table headers
        Range("I1").Value = "Ticker"
        Range("J1").Value = "Yearly Change"
        Range("K1").Value = "Percent Change"
        Range("L1").Value = "Total Stock Volume"
        
        'Loop through all stock transactions
        Dim i  As Long
        LastRow = WS.Cells(Rows.Count, 1).End(xlUp).Row
        For i = 2 To LastRow
            
            'Check if  still within the same stock ticker
            If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
                
                ' Write the unique Ticker symbol in the summary table
                Ticker_Symbol = Cells(i, 1).Value
                Range("I" & Summary_Table_Row).Value = Ticker_Symbol
                
                ' Calculate  and write the yearly change in the summary table
                Open_Price = Cells(2, 3).Value
                Close_Price = Cells(i, 6).Value
                Yearly_Change = Close_Price - Open_Price
                Range("J" & Summary_Table_Row).Value = Yearly_Change
                
                ' Calculate Ticker Percent_change and write in the summary table for each unique ticker
                If (Open_Price = 0 And Close_Price = 0) Then
                    Percent_Change = 0
                ElseIf (Open_Price = 0 And Close_Price <> 0) Then
                    Percent_Change = 1
                Else
                    Percent_Change = Yearly_Change / Open_Price
                    Cells(Summary_Table_Row, 11).Value = Percent_Change
                    Cells(Summary_Table_Row, 11).NumberFormat = "#.##"
                End If
                
                'Calculate Stock_Volume  and write in the summary table for each unique ticker
                Stock_Volume = Stock_Volume + Cells(i, 7).Value
                Cells(Summary_Table_Row, 12).Value = Stock_Volume
                Summary_Table_Row = Summary_Table_Row + 1
                Open_Price = Cells(i + 1, 3)
                Stock_Volume = 0
            Else
                Stock_Volume = Volume + Cells(i, 7).Value
            End If
        Next i
        
        'Conditional formatting
        LastRow1 = WS.Cells(Rows.Count, 9).End(xlUp).Row
        For j = 2 To LastRow1
            If Cells(j, 10).Value > 0 Then
                Cells(j, 10).Interior.Color = RGB(0, 100, 0)
            ElseIf Cells(j, 10).Value < 0 Then
                Cells(j, 10).Interior.Color = RGB(100, 0, 0)
            End If
        Next j
        
        'Challenges table labels
        Range("O2").Value = "Greatest % Increase"
        Range("O3").Value = "Greatest % Decrease"
        Range("O4").Value = "Greatest Total Volume"
        Range("P1").Value = "Ticker"
        Range("Q1").Value = "Value"
        
            'Challenges table calculation with VBA Application.WorksheetFunction
            For k = 2 To LastRow1
            If Cells(k, 11).Value = Application.WorksheetFunction.Max(WS.Range("K2:K" & LastRow1)) Then
                Cells(2, 16).Value = Cells(k, 9).Value
                Cells(2, 17).Value = Cells(k, 11).Value
                Cells(2, 17).NumberFormat = "#.##%"
            ElseIf Cells(k, 11).Value = Application.WorksheetFunction.Min(WS.Range("K2:K" & LastRow1)) Then
                Cells(3, 16).Value = Cells(k, 9).Value
                Cells(3, 17).Value = Cells(k, 11).Value
                Cells(3, 17).NumberFormat = "#.##%"
            ElseIf Cells(k, 12).Value = Application.WorksheetFunction.Max(WS.Range("L2:L" & LastRow1)) Then
                Cells(4, 16).Value = Cells(k, 9).Value
                Cells(4, 17).Value = Cells(k, 12).Value
            End If
        Next k
        
    Next WS
    
End Sub

