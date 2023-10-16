Attribute VB_Name = "Module1"
Sub Stock()

    Dim ws As Worksheet
    Dim LastRow As Long
    Dim SummaryRow As Long
    
    Dim Ticker As String
    Dim YearlyChange As Double
    Dim InitialValue As Double
    Dim FinalValue As Double
    Dim PercentChange As Double
    Dim TotalVolume As Double
        
    Dim GreatIncr As Double
    Dim TickerIncr As String
    Dim GreatDecr As Double
    Dim TickerDecr As String
    Dim GreatVol As Double
    Dim TickerVol As String
    
    
    'Running the code on all worksheets (don't forget to use ws. in front of all Cells and Range)
    
    For Each ws In ThisWorkbook.Worksheets
    
       'Setting initial values
        Ticker = ""
        YearlyChange = 0
        InitialValue = ws.Cells(2, 3).Value
        FinalValue = 0
        PercentChange = 0
        TotalVolume = 0
        SummaryRow = 2
        GreatIncr = 0
        TickerIncr = ""
        GreatDecr = 0
        TickerDecr = ""
        GreatVol = 0
        TickerVol = ""
    
        'Setting the Headers
    
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        
        'Setting Functionalities Headers
        
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
      
        'Finding the last not empty row to define the upper limit of the loop
        LastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
        
        For i = 2 To LastRow
    
            'Checking if we still withing the same stock
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            
                'Setting and Printing Ticker
                Ticker = ws.Cells(i, 1).Value
                ws.Cells(SummaryRow, 9).Value = Ticker
            
                'Setting Final Value
                FinalValue = ws.Cells(i, 6).Value
            
                'Calculating  and Printing Yearly Change
                YearlyChange = FinalValue - InitialValue
                
                ws.Cells(SummaryRow, 10).Value = YearlyChange
                ws.Cells(SummaryRow, 10).NumberFormat = "0.00"
                'Formating Cell
                    If YearlyChange < 0 Then
                        ws.Cells(SummaryRow, 10).Interior.ColorIndex = 3
                    Else
                        ws.Cells(SummaryRow, 10).Interior.ColorIndex = 4
                    End If
            
                'Calculating PercentChange
                PercentChange = (FinalValue / InitialValue) - 1
            
                'Set the percentage format for the cell - Used LINER GPT to help me to format like percentage
                ws.Cells(SummaryRow, 11).NumberFormat = "0.00%"
            
                'Printing PercentChange
                ws.Cells(SummaryRow, 11).Value = PercentChange
                    'Formating Cell
                    If PercentChange < 0 Then
                        ws.Cells(SummaryRow, 11).Interior.ColorIndex = 3
                    Else
                        ws.Cells(SummaryRow, 11).Interior.ColorIndex = 4
                    End If
            
                'Add to Total Stock Volume and Printing it
                TotalVolume = TotalVolume + ws.Cells(i, 7).Value
                ws.Cells(SummaryRow, 12).Value = TotalVolume
            
                'Add 1 to the Summary Row Counter
                SummaryRow = SummaryRow + 1
            
                'Reset the TotalVolume
                TotalVolume = 0
            
                'Defining the new InitialValue
                InitialValue = ws.Cells(i + 1, 3).Value
            
            Else
                'Add to Total Stock Volume
                TotalVolume = TotalVolume + ws.Cells(i, 7).Value
            End If
        
        
            'Testing Greatest Increases and Decreases
        
            If ws.Cells(i + 1, 11).Value > GreatIncr Then
                GreatIncr = ws.Cells(i + 1, 11).Value
                TickerIncr = ws.Cells(i + 1, 9).Value
            End If
        
            If ws.Cells(i + 1, 11).Value < GreatDecr Then
                GreatDecr = ws.Cells(i + 1, 11).Value
                TickerDecr = ws.Cells(i + 1, 9).Value
            End If
        
        
            'Testing Greatest Volume
            
            If ws.Cells(i + 1, 12).Value > GreatVol Then
                GreatVol = ws.Cells(i + 1, 12).Value
                TickerVol = ws.Cells(i + 1, 9).Value
            End If
                        
        Next i
        
        'Printing Functionality
        
        ws.Cells(2, 17).Value = GreatIncr
        ws.Cells(2, 17).NumberFormat = "0.00%"
        ws.Cells(2, 16).Value = TickerIncr
        
        ws.Cells(3, 17).Value = GreatDecr
        ws.Cells(3, 17).NumberFormat = "0.00%"
        ws.Cells(3, 16).Value = TickerDecr
        
        ws.Cells(4, 17).Value = GreatVol
        ws.Cells(4, 16).Value = TickerVol
    
    Next ws
    
End Sub




