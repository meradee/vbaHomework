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

Range("I1").Value = "Ticker"
Range("J1").Value = "Volume"

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
 Next ws
End Sub
