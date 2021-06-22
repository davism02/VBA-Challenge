Sub StockTicker()

For Each ws In Worksheets
'Dim ws As Worksheet
ws.Activate

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

' want to know how many tickers so we set up summary table

Dim SummaryTableRows As Integer

SummaryTableRows = 2


LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
MsgBox (ws.Name)
Cells(1, 9).Value = "Ticker"
Cells(1, 10).Value = "Yearly Change"
Cells(1, 11).Value = "Percent Change"
Cells(1, 12).Value = "Total Stock Volume"

For i = 2 To LastRow

    If Cells(i, 1).Value <> Cells(i + 1, 1).Value Then

    Ticker = Cells(i, 1).Value

    Cells(SummaryTableRows, 9).Value = Ticker
    
        
    TotalVolume = TotalVolume + Cells(i, 7).Value
    
    Cells(SummaryTableRows, 12).Value = TotalVolume
    
    TotalVolume = 0
    
    ClosePrice = Cells(i, 6).Value
    OpenPrice = Cells(i, 3).Value
    
    
    YearlyChange = ClosePrice - OpenPrice
     
     Cells(SummaryTableRows, 10).Value = YearlyChange
     
   'This code added so we don't divide by zero if open price is zero
    If OpenPrice > 0 Then
    PercentChange = (YearlyChange / OpenPrice)
    End If
    
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

Cells(2, 15).Value = "Greatest % Increase"
Cells(3, 15).Value = "Greatest % Decreae"
Cells(4, 15).Value = "Greatest Total Volume"
Cells(1, 16).Value = "Ticker"
Cells(1, 17).Value = "Value"


'Find Max Values

Dim vMin, vMax, totallMax
 vMax = Application.WorksheetFunction.Max(Columns("J"))
 Range("Q2").Value = vMax
 
 vMin = Application.WorksheetFunction.Min(Columns("J"))
 Range("Q3").Value = vMin
 
 TotalMax = Application.WorksheetFunction.Max(Columns("L"))
 Range("Q4").Value = TotalMax

 'Loop Through to Find Stock ticker for max percent increase
    
    For i = 1 To LastRow

        
        If Cells(i, 10).Value = Application.WorksheetFunction.Max(Columns("J")) Then
            
            Cells(2, 16).Value = Cells(i, 9)
            
            ' If first match is found, exit the for loop
            Exit For
        
        End If
    
    
    Next i
 
 'Loop Through to Find Stock ticker for min percent increase
    
    For i = 1 To LastRow

        
        If Cells(i, 10).Value = Application.WorksheetFunction.Min(Columns("J")) Then
            
            Cells(3, 16).Value = Cells(i, 9)
            
            ' If first match is found, exit the for loop
            Exit For
        
        End If
    Next i


'Loop Through to Find Stock ticker for most volume

    For i = 1 To LastRow

        
        If Cells(i, 12).Value = Application.WorksheetFunction.Max(Columns("L")) Then
            
            Cells(4, 16).Value = Cells(i, 9)
            
            ' If first match is found, exit the for loop
            Exit For
        
        End If
    Next i

Next ws

End Sub
