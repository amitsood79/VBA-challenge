Sub market_stock()

'Variable Declaration
Dim voltotal As Double
Dim open_price As Double
Dim close_price As Double
Dim ws As Worksheet
Dim top As Long
Dim bottom As Long
Dim yearly_change As Double
Dim percent_change As Double
Dim greatest_increase As Double
Dim greatest_decrease As Double
Dim greatest_total_vol As Double

'Variable Assignment
voltotal = 0
greatest_increase = 0
greatest_decrease = 0
greatest_total_vol = 0

'Looping through each sheet
For Each ws In Worksheets

'Creating new columns
ws.Range("I1").EntireColumn.Insert
ws.Range("J1").EntireColumn.Insert
ws.Range("K1").EntireColumn.Insert
ws.Range("L1").EntireColumn.Insert
ws.Range("P1").EntireColumn.Insert
ws.Range("Q1").EntireColumn.Insert

'Creating column heaaders
ws.Range("I1").Value = "Ticker"
ws.Range("J1").Value = "Yearly Change"
ws.Range("K1").Value = "Percent Change"
ws.Range("L1").Value = "Total Stock Volume"
ws.Range("P1").Value = "Ticker"
ws.Range("Q1").Value = "Value"

'Autofitting the columns
ws.Columns("J:L").AutoFit
ws.Columns("O").ColumnWidth = 25
ws.Columns("P").ColumnWidth = 10
ws.Columns("Q").ColumnWidth = 15

'Finding Last Rows
Lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
Lastrow_vol = ws.Cells(Rows.Count, 9).End(xlUp).Row
Lastrow_vol = 2

'Looping through columns
  For x = 2 To Lastrow
    
    If ws.Cells(x + 1, 1).Value <> ws.Cells(x, 1).Value Then
        
        ticker = ws.Cells(x, 1).Value
        
        voltotal = voltotal + ws.Cells(x, 7).Value
        
        top = ws.Range("A:A").find(what:=ws.Cells(x, 1).Value, lookat:=xlWhole).Row

        bottom = ws.Range("A:A").find(what:=ws.Cells(x, 1).Value, searchdirection:=xlPrevious, lookat:=xlWhole).Row

        open_price = ws.Cells(top, 3).Value
        
        close_price = ws.Cells(bottom, 6).Value
        
        yearly_change = Round((close_price - open_price), 2)
        
        percent_change = (yearly_change / open_price)
                      
        ws.Range("I" & Lastrow_vol).Value = ticker
        
        ws.Range("L" & Lastrow_vol).Value = voltotal
        
        'Finding the greatest total volume
            If ws.Range("L" & Lastrow_vol).Value >= greatest_total_vol Then
                greatest_total_vol = ws.Range("L" & Lastrow_vol).Value
                ws.Range("O4").Value = "Greatest Total Volume"
                ws.Range("P4").Value = ws.Range("I" & Lastrow_vol).Value
                ws.Range("Q4").Value = greatest_total_vol
            End If
            
        ws.Range("J" & Lastrow_vol).Value = yearly_change
        
        'Conditional Formatting
            If yearly_change < 0 Then
            ws.Range("J" & Lastrow_vol).Interior.ColorIndex = 3
            Else
            ws.Range("J" & Lastrow_vol).Interior.ColorIndex = 4
            End If
        
        ws.Range("K" & Lastrow_vol).Value = percent_change
        ws.Range("K:K").NumberFormat = "0.00%"
                
        'Finding greatest increase
            If ws.Range("K" & Lastrow_vol).Value >= greatest_increase Then
                greatest_increase = ws.Range("K" & Lastrow_vol).Value
                ws.Range("O2").Value = "Greatest % Increase"
                ws.Range("P2").Value = ws.Range("I" & Lastrow_vol).Value
                ws.Range("Q2").Value = greatest_increase
                ws.Range("Q2").NumberFormat = "0.00%"
            End If
            
        'Finding greatest decrease
            If ws.Range("K" & Lastrow_vol).Value <= greatest_decrease Then
                greatest_decrease = ws.Range("K" & Lastrow_vol).Value
                ws.Range("O3").Value = "Greatest % Decrease"
                ws.Range("P3").Value = ws.Range("I" & Lastrow_vol).Value
                ws.Range("Q3").Value = greatest_decrease
                ws.Range("Q3").NumberFormat = "0.00%"
            End If
                                                                                        
        Lastrow_vol = Lastrow_vol + 1
        
        voltotal = 0
        
        open_price = 0
        
        close_price = 0
        
        yearly_change = 0
                     
    Else
                  
        voltotal = voltotal + ws.Cells(x, 7).Value
        
    End If
          
               
    Next x

greatest_increase = 0
greatest_decrease = 0
greatest_total_vol = 0

Next ws

End Sub
