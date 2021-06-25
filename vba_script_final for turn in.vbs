Sub Stock()
' Declare Variables
    Dim ws As Worksheet
    Dim Ticker As String
    Dim Open_Price As Double
    Dim Close_Price As Double
    Dim Year_Change As Double
    Dim Total_Volume As Double
    Dim Percent_Change As Double
    Dim Ticker_Increase As String
    Dim Ticker_Decrease As String
    Dim Ticker_Greatest As String
    Dim Increase_Value As Double
    Dim Decrease_Value As Double
    Dim Greatest_Value As Double
    Dim LastRow As Long ' find the last row in column A  for Ticker loop
    Dim LaswRow2 As Long  ' Find last row in the summary table column I
    Dim NxRow As Long
   
      
For Each ws In Worksheets

' Set initial variables
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        Total_Volume = 0
        Open_Price = ws.Range("C2")
        NxRow = 1

' Add Headers
        ws.Range("I1") = "Ticker"
        ws.Range("j1") = "Yearly Change"
        ws.Range("K1") = "Percent Change"
        ws.Range("L1") = "Total Stock Volume"

        ws.Range("P1") = "Ticker"
        ws.Range("Q1") = "Value"
        ws.Range("O2") = "Greatest PCT Increase"
        ws.Range("O3") = "Greatest PCT Decrease"
        ws.Range("O4") = "Greatest Total Volume"

' Loop through ticker to determine if next cell is equal to current cell if so we cam sum the values for each unique ticker..

For i = 2 To LastRow
'If next row ticker equals the currrent ticker symbol add to total volume
            If ws.Cells(i, 1) = ws.Cells(i + 1, 1) Then
                Total_Volume = Total_Volume + ws.Cells(i, 7)

    'If next ticker is not equal to curren then you are at the last ticker so you know the closing pirce.  
        
            Else
                Close_Price = ws.Cells(i, 6)
                Total_Volume = Total_Volume + ws.Cells(i, 7)
                Year_Change = Close_Price - Open_Price
                On Error Resume Next  ' If Open_Price is 0 there will be an error
                Percent_Change = (Close_Price - Open_Price) / Open_Price
                If Err.Number <> 0 Then
                    Percent_Change = 0
                End If
                NxRow = NxRow + 1
                ws.Cells(NxRow, 9) = ws.Cells(i, 1)
                ws.Cells(NxRow, 10) = Year_Change
                ws.Cells(NxRow, 11) = Percent_Change
                ws.Cells(NxRow, 12) = Total_Volume
                Total_Volume = 0
                Open_Price = ws.Cells(i + 1, 3)
            End If
        Next i
'We now have summary table so we will loop through colunm I so we will need the last row
        LastRow2 = ws.Cells(Rows.Count, 9).End(xlUp).Row

' Variables for the max, min and greatest volume. Starting at the first ticker in I2 will loop throgh to compare each cell for greatest increse, decrase and max volume
        Ticker_Increase = ws.Range("I2")
        Ticker_Decrease = ws.Range("I2")
        Ticker_Greatest = ws.Range("I2")
        Increase_Value = ws.Range("K2")
        Decrease_Value = ws.Range("K2")
        Greatest_Value = ws.Range("L2")

' For Loop to Find Greatest Increase
        For j = 2 To LastRow2
            If ws.Cells(j + 1, 11) > Increase_Value And _
            ws.Cells(j + 1, 11) <> "" Then
                Increase_Value = ws.Cells(j + 1, 11)
                Ticker_Increase = ws.Cells(j + 1, 9)
            End If
        Next j
' Loop to Find Greatest Decrease
        For j = 2 To LastRow2
            If ws.Cells(j + 1, 11) < Decrease_Value And _
            ws.Cells(j + 1, 11) <> "" Then
                Decrease_Value = ws.Cells(j + 1, 11)
                Ticker_Decrease = ws.Cells(j + 1, 9)
            End If
        Next j
' Loop to Find Greatest Total Volume
        For j = 2 To LastRow2
            If ws.Cells(j + 1, 12) > Greatest_Value And _
            ws.Cells(j + 1, 12) <> 0 Then
                Greatest_Value = ws.Cells(j + 1, 12)
                Ticker_Greatest = ws.Cells(j + 1, 9)
            End If
        Next j

' Add values to the proper cells on sheet
        ws.Range("P2") = Ticker_Increase
        ws.Range("P3") = Ticker_Decrease
        ws.Range("P4") = Ticker_Greatest
        ws.Range("Q2") = Increase_Value
        ws.Range("Q3") = Decrease_Value
        ws.Range("Q4") = Greatest_Value

' If balue is less than 0 code to red.  If greater than zero code to greeen
        For j = 2 To LastRow2
            If ws.Cells(j, 10) <= 0 Then
                ws.Cells(j, 10).Interior.Color = RGB(255, 0, 0)
            Else
                ws.Cells(j, 10).Interior.Color = RGB(0, 255, 0)
            End If
        Next j

'   Format cells
        ws.Range("j2:j" & LastRow2).NumberFormat = "0.00"
        ws.Range("K2:K" & LastRow2).NumberFormat = "0.00%"
        ws.Range("Q2:Q3").NumberFormat = "0.00%"
        ws.Cells.Columns.AutoFit

    Next ws

'

End Sub

