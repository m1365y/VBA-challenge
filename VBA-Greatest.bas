Attribute VB_Name = "Module3"
Sub StockAnalysis()
    Dim ws As Worksheet
For Each ws In ActiveWorkbook.Worksheets
    
    Dim lastRow As Long
    Dim ticker As String
    Dim openPrice As Double
    Dim closePrice As Double
    Dim yearlyChange As Double
    Dim percentChange As Double
    Dim totalVolume As Double
    Dim SummaryRow As Long
    Dim i As Long
 '-----------------------------------------
    Dim maxPercentIncreaseTicker As String
    Dim maxPercentDecreaseTicker As String
    Dim maxTotalVolumeTicker As String
    Dim maxPercentIncrease As Double
    Dim maxPercentDecrease As Double
    Dim maxTotalVolume As Double
    
    
    
    'Find last row of data in column A
    lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    '--------------------------------------------
    
    'Initialize summary table row
    SummaryRow = 2
    
    'Initialize total volume counter
    totalVolume = 0
    
    'Loop through all rows of data
    For i = 2 To lastRow
    
        'Check if ticker symbol has changed
        If ws.Cells(i, 1).Value <> ws.Cells(i - 1, 1).Value Then
            
            'Set ticker symbol
            ticker = ws.Cells(i, 1).Value
            
            'Set opening price
            openPrice = ws.Cells(i, 3).Value
            
        End If
        
        'Add volume to total volume
        totalVolume = totalVolume + ws.Cells(i, 7).Value
        
        'Check if current row is last row for current ticker symbol
        If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
            
            'Set closing price
            closePrice = ws.Cells(i, 6).Value
            
            'Calculate yearly change and percent change
            yearlyChange = closePrice - openPrice
            percentChange = yearlyChange / openPrice
            
            '-----------------------------
            
            'Reset total volume counter
            totalVolume = 0
            
            'Move to next row of summary table
            SummaryRow = SummaryRow + 1
            
        End If
        
        
        '--------------------------------------------
      
    
    If percentChange > maxPercentIncrease Then
                maxPercentIncrease = percentChange
                maxPercentIncreaseTicker = ticker
            ElseIf percentChange < maxPercentDecrease Then
                maxPercentDecrease = percentChange
                maxPercentDecreaseTicker = ticker
            End If
            
            If totalVolume > maxTotalVolume Then
                maxTotalVolume = totalVolume
                maxTotalVolumeTicker = ticker
            End If
        
        
        
        
    Next i
    
    '----------------------------------------------------------
    ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"
        
        
        
        ws.Range("o2").Value = "Greatest % Increase"
        ws.Range("p2").Value = maxPercentIncreaseTicker
        ws.Range("Q2").Value = maxPercentIncrease
        ws.Range("Q2").NumberFormat = "0.00%"
        
        
        ws.Range("o3").Value = "Greatest % Decrease"
        ws.Range("p3").Value = maxPercentDecreaseTicker
        ws.Range("Q3").Value = maxPercentDecrease
        ws.Range("Q3").NumberFormat = "0.00%"
        
        
        
        ws.Range("o4").Value = "Greatest Total Volume"
        ws.Range("p4").Value = maxTotalVolumeTicker
        ws.Range("Q4").Value = maxTotalVolume
        ws.Range("Q4").NumberFormat = "0.00%"
            
            
    
    Next
    
End Sub


