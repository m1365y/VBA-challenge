Attribute VB_Name = "Module2"
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
 
    
    
    
    'Find last row of data in column A
    lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    'Set up summary table headers
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"
    
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
            
            'Write summary data to summary table
            ws.Range("I" & SummaryRow).Value = ticker
            ws.Range("J" & SummaryRow).Value = yearlyChange
            ws.Range("K" & SummaryRow).Value = percentChange
            ws.Range("K" & SummaryRow).NumberFormat = "0.00%"
            ws.Range("L" & SummaryRow).Value = totalVolume
            
            'Reset total volume counter
            totalVolume = 0
            
            'Move to next row of summary table
            SummaryRow = SummaryRow + 1
            
        End If
        
        
     
        
        
        
    Next i
    
            
    
    Next
    
End Sub


