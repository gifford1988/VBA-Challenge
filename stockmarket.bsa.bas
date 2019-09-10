Attribute VB_Name = "Module1"
Sub total()

' For Each was added to the syntax so that all of the elements of the sheets are added.

For Each WS In Worksheets

' The variables are listed out.  Variant is used because it can hold numbers up to a double.  String is used because  it is able to hold characters.

Dim ticker As String
Dim counter As Variant
Dim total As Variant
Dim year_open As Variant
Dim year_close As Variant
Dim last_row As Variant
Dim last_ticker As Variant
Dim yearly_change As Variant
Dim percent_change As Variant

' Titles of the columns are added for the worksheets.  Adding WS ahead of every range automatically inputs the value into all of the worksheets at the location specified.

WS.Range("I1").Value = "Ticker"
WS.Range("J1").Value = "Yearly Change"
WS.Range("K1").Value = "Percent Change"
 WS.Range("L1").Value = "Total Stock Volume"
    
' Adding the equation, ws.cells(rows.count,1).end(xlup).row, will run a loop from the first number to the last of a column.

last_row = WS.Cells(Rows.Count, 1).End(xlUp).Row

' The variables are next given values.  A ticker is defined as, " "', to  hold characters.  Counter is defined as one because it counts one of a data point and/or counter plus one.  Year_open and close are zero due to the data points starting at zero.
     
ticker = “”
counter = 1
total = 0
year_open = 0
year_close = 0
        


' Starts at two then goes to the last row of the column 1 based on what last_row was defined as.

For Row = 2 To last_row:
        
' Ifthen statement next added.  It states if row two in column one is not equal to a ticker, then the counter goes to counter plus one while looking for a new name.
    
If WS.Cells(Row, 1).Value <> ticker Then
counter = counter + 1
                
' Sets the original/current ticker to a new name based on the new ticker.

ticker = WS.Cells(Row, 1).Value
                
' Sets the year_open to the opening number of the first data point of the column.

year_open = WS.Cells(Row, 3).Value
                    
' The ticker names are set in a new row for the ticker column.

WS.Cells(counter, 9).Value = ticker
                
' Sets the beginning total to the first set of volumes recorded .

total = WS.Cells(Row, 7).Value
WS.Cells(counter, 12) = total
            
Else

' If ticker is the same or equal, add its volume to the total.  l

total = total + WS.Cells(Row, 7).Value
WS.Cells(counter, 12).Value = total

' End if the last entry for the ticker has been found.

End If

If WS.Cells((Row + 1), 1).Value <> ticker Then
            


' This sets the value of year_close
       
year_close = WS.Cells(Row, 6).Value

' Calculates the net increase or decrease from the year open to close.

 yearly_change = year_close - year_open

' This adds the calculation for yearly_change to the Yearly Change column
                    
WS.Cells(counter, 10).Value = yearly_change

' Calculates the yearly_change into a percentage.

If year_open = 0 Then
percent_change = yearly_change
Else
percent_change = yearly_change / year_open
End If
                    
' Applies the yearly change to the Percent Change Column and reformats to a percentage.

WS.Cells(counter, 11).Value = percent_change
WS.Cells(counter, 11).NumberFormat = "0.00%"
            
' Cells are formatted, if greater than zero then green.
 
If yearly_change > 0 Then
WS.Cells(counter, 10).Interior.Color = vbGreen

' Cells are formatted, if less than zero then red.

ElseIf yearly_change < 0 Then
WS.Cells(counter, 10).Interior.Color = vbRed
                
End If
            
End If
            
Next Row

' Variables are defined using variant as it holds both integers and doubles.

Dim top_increase As Variant
Dim top_decrease As Variant
Dim top_volume As Variant

' Title input for both rows and columns.
        
WS.Range("O2").Value = "Greatest % Increase"
WS.Range("O3").Value = "Greatest % Decrease"
WS.Range("O4").Value = "Greatest Total Volume"
WS.Range("P1").Value = "Ticker"
WS.Range("Q1").Value = "Value"
        
' This formula finds the last row of the column list.

last_ticker = WS.Cells(Rows.Count, 9).End(xlUp).Row

' Variables are set at zero due to the numbers starting at zero.

top_increase = 0
top_decrease = 0
top_total = 0
        
For Row = 2 To last_ticker:
            
' If the Positive Change is higher than the current leader.

If WS.Cells(Row, 11).Value > top_increase Then
                               
' Current leader is set.

top_increase = WS.Cells(Row, 11).Value
                
' Ticker and Percent are entered into the Leaderboard.

WS.Range("P2").Value = WS.Cells(Row, 9).Value
WS.Range("Q2").Value = WS.Cells(Row, 11).Value
            
            End If
            
' If the negative change is lower than the current leader.

If WS.Cells(Row, 11).Value < top_decrease Then
                
' New Current Leader is set.

top_decrease = WS.Cells(Row, 11).Value
            

' Ticker and its percentage are entered into the leaderboard.

WS.Range("P3").Value = WS.Cells(Row, 9).Value
WS.Range("P3").NumberFormat = "0.00%"
WS.Range("Q3").Value = WS.Cells(Row, 11).Value
WS.Range("Q3").NumberFormat = "0.00%"
            
End If
            
' If statement reflecting if total stock volume is higher than current highest leader.

If WS.Cells(Row, 12).Value > top_total Then
                
' New Current Leader is set

top_total = WS.Cells(Row, 12).Value

' Enter the total leader for percentage and ticker.
             
WS.Range("P4").Value = WS.Cells(Row, 9).Value
WS.Range("Q4").Value = WS.Cells(Row, 11).Value
            
End If
        
Next Row
        
Next WS
    
End Sub

