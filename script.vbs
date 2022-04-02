Attribute VB_Name = "Module1"
Sub Analysis()
    ' BONUS: Variable holds number of sheets
    Dim Sheet_Count As Integer
    
    Sheet_Count = ThisWorkbook.Worksheets.Count
    
    For j = 1 To Sheet_Count
        
        ' BONUS: Selects current worksheet
        Worksheets(j).Select
        
        ' Headers
        Cells(1, 9).Value = "Ticker"
        Cells(1, 10).Value = "Yearly Change"
        Cells(1, 11).Value = "Percent Change"
        Cells(1, 12).Value = "Total Stock Volume"
        
        ' Variable holds the ticker symbol
        Dim Ticker_Symbol As String
        
        ' Tracks opening value of the year and closing value of the year
        Dim Opening_Value As Double
        Opening_Value = Cells(2, 3).Value
        Dim Closing_Value As Double
        
        ' Variable holds yearly change
        Dim Yearly_Change As Double
        
        ' Variable holds percent change
        Dim Percent_Change As Double
        
        ' Variable holds the total stock volume
        Dim Total_Stock_Volume As Double
        Total_Stock_Volume = 0
        
        ' Tracks the location for each stock in the output table
        Dim Ticker_Symbol_Row As Integer
        Ticker_Symbol_Row = 2
        
        ' BONUS: Variables to hold the Greatest % Increase, Greatest % Decrease, and Greatest Total Volume
        Dim Greatest_Percent_Increase As Double
        Greatest_Percent_Increase = 0
        Dim Greatest_Percent_Decrease As Double
        Greatest_Percent_Decrease = 0
        Dim Greatest_Total_Volume As Double
        Greatest_Total_Volume = 0
        
        ' BONUS: Variables to hold previous ticker symbols
        Dim GPI As String
        Dim GPD As String
        Dim GTV As String
        
        ' Loop through all rows of stock's opening and closing information
        For i = 2 To ActiveSheet.UsedRange.Rows.Count
        
            ' Check if we are not still within the same stock
            If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
            
                ' Set the ticker symbol
                Ticker_Symbol = Cells(i, 1).Value
                
                
                ' Calculates yearly change
                Closing_Value = Cells(i, 6).Value
                Yearly_Change = Closing_Value - Opening_Value
                
                'Calculated percent change
                Percent_Change = (Yearly_Change / Opening_Value) * 100
                
                ' BONUS: Checking for Greatest % Increase and % Decrease
                If Percent_Change > Greatest_Percent_Increase Then
                    GPI = Ticker_Symbol
                    Greatest_Percent_Increase = Percent_Change
                End If
                
                If Percent_Change <= Greatest_Percent_Decrease Then
                    GPD = Ticker_Symbol
                    Greatest_Percent_Decrease = Percent_Change
                End If
                
                
                ' Add to the total stock volume
                Total_Stock_Volume = Total_Stock_Volume + Cells(i, 7).Value
                
                ' BONUS: Checking for Greatest Total Volume
                If Total_Stock_Volume >= Greatest_Total_Volume Then
                    GTV = Ticker_Symbol
                    Greatest_Total_Volume = Total_Stock_Volume
                End If
                
                ' Print the ticker symbol to the output table
                Range("I" & Ticker_Symbol_Row).Value = Ticker_Symbol
                
                ' Print yearly change to the output table
                Range("J" & Ticker_Symbol_Row).Value = Yearly_Change
                
                'Print percent change to the output table
                Range("K" & Ticker_Symbol_Row).Value = Percent_Change
                
                ' Print the total stock volume to the output table
                Range("L" & Ticker_Symbol_Row).Value = Total_Stock_Volume
                
                ' Formatting yearly change color
                If Range("J" & Ticker_Symbol_Row).Value >= 0 Then
                    
                    Range("J" & Ticker_Symbol_Row).Interior.ColorIndex = 4
                
                Else
                    
                    Range("J" & Ticker_Symbol_Row).Interior.ColorIndex = 3
                    
                End If
                    
                ' Formatting yearly change decimal
                Range("J" & Ticker_Symbol_Row).NumberFormat = "0.00"
                
                ' Formatting percent change decimal and percent symbol
                Range("K" & Ticker_Symbol_Row).NumberFormat = "0.00\%"
                
                ' Add one to the output table row
                Ticker_Symbol_Row = Ticker_Symbol_Row + 1
                
                ' New stock opening value
                Opening_Value = Cells(i + 1, 3).Value
                
                ' Reset the total stock volume
                Total_Stock_Volume = 0
            
            ' If the cell immediately following a row is the same stock
            Else
            
                ' Add to total stock volume
                Total_Stock_Volume = Total_Stock_Volume + Cells(i, 7).Value
            
            End If
            
        Next i
        
        ' BONUS: Printing bonus table row headers
        Range("O" & 2).Value = "Greatest % Increase"
        Range("O" & 3).Value = "Greatest % Decrease"
        Range("O" & 4).Value = "Greatest Total Volume"
        
        ' BONUS: Printing bonus table column headers
        Range("P" & 1).Value = "Ticker"
        Range("Q" & 1).Value = "Value"
        
        ' BONUS: Printing bonus table tickers and values
        Range("P" & 2).Value = GPI
        Range("P" & 3).Value = GPD
        Range("P" & 4).Value = GTV
        
        Range("Q" & 2).Value = Greatest_Percent_Increase
        Range("Q" & 2).NumberFormat = "0.00\%"
        Range("Q" & 3).Value = Greatest_Percent_Decrease
        Range("Q" & 3).NumberFormat = "0.00\%"
        Range("Q" & 4).Value = Greatest_Total_Volume
        
        
    Next j

End Sub


