Attribute VB_Name = "Module1"
Sub stock_homework()

'Runs loop code through each worksheet
For Each ws In Worksheets

'Define variables
Dim ticker_name As String
Dim total_volume As Double

total_volume = 0

'Counter starting at 2nd row
Dim summary_count As Integer
summary_count = 2

'Last row of the data worksheet
FinalRow = Cells(Rows.Count, 1).End(xlUp).Row

'Naming
ws.Range("I1").Value = "Ticker"
ws.Range("J1").Value = "Total Stock Volume"

'For loop
For i = 2 To FinalRow

If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then

'1. Ticker for ith row
ticker_name = ws.Cells(i, 1).Value
'2. Volume for the ith row
total_volume = total_volume + ws.Cells(i, "G").Value

'3. Print ticker in cells
ws.Range("I" & summary_count).Value = ticker_name
ws.Range("J" & summary_count).Value = total_volume

'4. Increase to next row
summary_count = summary_count + 1
'5. Recount volume back to 0 for the next total volume count
total_volume = 0

Else
'If stock ticker same ith and ith+1 row, keep adding volume together
 total_volume = total_volume + ws.Cells(i, "G").Value
 
End If

 Next i
 

'Define Variables ----------------------Second Part------------------------
Dim opening_price As Double
Dim closing_price As Double

Dim yearly_change As Double
Dim percent_change As Double

'Row start on 2
summary_count = 2

'Cell of opening price begins at here
opening_price = ws.Range("C2").Value

'Naming
ws.Range("K1").Value = "Yearly Change"
ws.Range("L1").Value = "Percent Change"

'For loop finding yearly change and percent change
For i = 3 To FinalRow

If (Cells(i, 1).Value <> Cells(i + 1, 1).Value) Then

'Define closing price value
closing_price = Cells(i, 6).Value

'Calculate yearly change
yearly_change = closing_price - opening_price

'Calculate percent change
If opening_price <> 0 Then
 percent_change = yearly_change / opening_price
Else
 percent_change = 0
End If

'Format percent change to percentage
ws.Range("L" & summary_count).NumberFormat = "0.00%"

'Prints in cell
ws.Range("K" & summary_count).Value = yearly_change
ws.Range("L" & summary_count).Value = percent_change

'Colors red and green for yearly change
If yearly_change >= 0 Then
ws.Range("K" & summary_count).Interior.ColorIndex = 4
Else
ws.Range("K" & summary_count).Interior.ColorIndex = 3
End If

summary_count = summary_count + 1

End If
 Next i

 Next ws
End Sub
