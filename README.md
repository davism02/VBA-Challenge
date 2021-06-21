Sub StockTicker()


'Identify unique Ticker name

Dim Ticker As String
Dim OpenPrice As Double
Dim ClosePrice As Double
Dim YearlyChange As Double
Dim TotalVolume As Double
Dim PercentChange As Double
Dim WorksheetName As String
TotalVolume = 0

Dim LastRow As Long



 
' want to know how many tickers  so we set up summary table

Dim SummaryTableRows As Integer

SummaryTableRows = 2



For Each ws In Worksheets

LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

    Cells(1, 9).Value = "Ticker"
    Cells(1, 10).Value = "Yearly Change"
    Cells(1, 11).Value = "Percent Change"
    Cells(1, 12).Value = "Total Stock Volume"

    For i = 2 To LastRow
    
        If Cells(i, 1).Value <> Cells(i + 1, 1).Value Then
    
        Ticker = Cells(i, 1).Value
    
        Cells(SummaryTableRows, 9).Value = Ticker
        
        OpenPrice = Cells(i, 3).Value
        
        TotalVolume = TotalVolume + Cells(i, 7).Value
        
        Cells(SummaryTableRows, 12).Value = TotalVolume
        
        TotalVolume = 0
        
        ClosePrice = Cells(i, 6).Value
        
        YearlyChange = ClosePrice - OpenPrice
         
         Cells(SummaryTableRows, 10).Value = YearlyChange
         
                      
        
        PercentChange = (YearlyChange / OpenPrice)
        
        Cells(SummaryTableRows, 11).Value = PercentChange
        
     
            'Format Yearly change
            
               
                   If YearlyChange < 0 Then
                      Cells(SummaryTableRows, 10).Interior.ColorIndex = 3
                      
                                         
                   ElseIf YearlyChange >= 0 Then
                   Cells(SummaryTableRows, 10).Interior.ColorIndex = 4
                    End If
      
                           
                           'Format Percentage change
                     Cells(SummaryTableRows, 11).NumberFormat = "0.00%"
                              
                          
      'Add 1 to Summary table to go to next row
      
        SummaryTableRows = SummaryTableRows + 1
          
    Else: TotalVolume = TotalVolume + Cells(i, 12).Value
        
            Cells(i, 12).NumberFormat = "General"
        
        
    End If
            
 Next i

    Cells(2, 15).Value = "Greatess % Increase"
    Cells(3, 15).Value = "Greatest % Decreae"
    Cells(4, 15).Value = "Greates Total Volume"
    Cells(1, 16).Value = "Ticker"
    Cells(1, 17).Value = "Value"
            

Next ws

        
End Sub


               
    

