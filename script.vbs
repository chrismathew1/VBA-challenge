Attribute VB_Name = "Module1"
Sub Ticker()
    ' Variable holds number of sheets
    Dim Sheet_Count As Integer
    
    Sheet_Count = ThisWorkbook.Worksheets.Count
    
    For j = 1 To Sheet_Count
        
        ' Selects current worksheet
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
                
                ' Add to the total stock volume
                Total_Stock_Volume = Total_Stock_Volume + Cells(i, 7).Value
                
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
    Next j

End Sub


