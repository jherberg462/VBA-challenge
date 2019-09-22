Sub stock_data():



'create variable for running volume total, opening and closing prices
Dim volume As Long
Dim closing_price As Double
Dim open_price As Double

'create variable for percent and yearly change
Dim percent_change As Double
Dim yearly_change As Double

'create counter for number of stocks analyized, starting at 2 for placement on summary table
Dim ticker_no As Integer



'count number of worksheets
Dim worksheets As Integer
Dim i As Integer
worksheets = ActiveWorkbook.worksheets.Count
For i = 1 To worksheets  'will uncomment once code for one sheet is complete

ticker_no = 2 ' set ticker_no to 2 inside of worksheet loop

'setup headers
Range("K1").Value = "Ticker"
Range("L1").Value = "Yearly Change"
Range("M1").Value = "Percent Change"
Range("N1").Value = "Total Stock Volume"


'count number of rows
Dim last_row As Long
last_row = Cells(Rows.Count, 1).End(xlUp).row

'loop through each row in ws
Dim r As Long
For r = 2 To last_row



'if statement for when there is a new ticker symbol
If Cells(r - 1, 1).Value <> Cells(r, 1) Then
'save opening price for ticker
open_price = Cells(r, 3).Value

'add first days volume to running total of volume
'volume = Cells(r, 7).Value
Cells(ticker_no, 14).Value = Cells(r, 7).Value

'check if next row uses the same ticker symbol
ElseIf Cells(r, 1).Value = Cells(r + 1, 1).Value Then

'add days volume to running total
'volume = volume + Cells(r, 7).Value
Cells(ticker_no, 14).Value = Cells(ticker_no, 14).Value + Cells(r, 7).Value

'if statement for when the last row for the ticker symbol has being tabulated
ElseIf Cells(r, 1).Value <> Cells(r + 1, 1).Value Then

'add last days volume to running total
Cells(ticker_no, 14).Value = Cells(ticker_no, 14).Value + Cells(r, 7).Value

'print total volume
'Cells(ticker_no, 14).Value = volume

'reset volume back to zero
'volume = 0 'this is now unnecessary due to removing this counter for volume

'find closing price
closing_price = Cells(r, 6)

'find yearly change, and print value
yearly_change = closing_price - open_price
Cells(ticker_no, 12).Value = yearly_change

'find percenge change and print value
if open_price <> 0 then 'if statement to prevent dividing by zero
percent_change = (yearly_change / open_price)
Cells(ticker_no, 13).Value = percent_change
Else
cells(ticker_no, 13).Value = 0
end if 

Cells(ticker_no, 13).NumberFormat = "0.000%" 'format to percentage


'find and print ticker symbol
Cells(ticker_no, 11).Value = Cells(r, 1).Value

'if statement for conditional formatting
If yearly_change > 0 Then
Cells(ticker_no, 12).Interior.ColorIndex = 4 'set to green for positive return

ElseIf yearly_change < 0 Then
Cells(ticker_no, 12).Interior.ColorIndex = 3 'set to red for negative return
End If
'end of conditional formatting

'add one to counter for summary table
ticker_no = ticker_no + 1
End If

Next r


'get top volume, and best/worse performance

'count number of rows in summary section
last_row = Cells(Rows.Count, 11).End(xlUp).row

'setup headers
Cells(1, 18).Value = "Ticker"
Cells(1, 19).Value = "Value"
Cells(2, 17).Value = "Greatest % Increase"
Cells(3, 17).Value = "Greatest % Decrease"
Cells(4, 17).Value = "Highest Volume"


'set initial performance to zero for top performers
Range("S2:S4").Value = 0

'create loop to check performance
For r = 2 To last_row

'calculate best performance
If Cells(r, 13).Value > Cells(2, 19).Value Then 'dont need to confirm above zero because max value starts at zero
Cells(2, 19).Value = Cells(r, 13).Value
Cells(2, 18).Value = Cells(r, 11).Value '11 is the column index for the ticker
End If

'calculate worse performance
If Cells(r, 13).Value < Cells(3, 19).Value Then 'dont need to confirm below zero because min value starts at zero
Cells(3, 19).Value = Cells(r, 13).Value
Cells(3, 18).Value = Cells(r, 11).Value '11 is the column index for the ticker
End If

'calculate highest volume
If Cells(r, 14).Value > Cells(4, 19).Value Then 'dont need to confirm above zero because max value starts at zero
Cells(4, 19).Value = Cells(r, 14).Value
Cells(4, 18).Value = Cells(r, 11).Value '11 is the column index for the ticker
End If

Next r
'format best and worse performance cells to percentage
Cells(3, 19).NumberFormat = "0.00%"
Cells(2, 19).NumberFormat = "0.00%"
Cells(4, 19).NumberFormat = "General"

'cleaning up formatting
Cells(1, 11).Columns.AutoFit 'resizing columns so their width is appropiate according to header cell
Cells(1, 12).Columns.AutoFit
Cells(1, 13).Columns.AutoFit
Cells(1, 14).Columns.AutoFit
Cells(3, 17).Columns.AutoFit

If i < worksheets Then
'select next worksheets
ActiveSheet.Next.Activate
End If



Next i


MsgBox ("macro is done")
'need to figure out overflow error with running total of volume in the middle if statement (is currently commented out)
'fixed volume issue in a hacky way -- need to find out if there is a better way to do this
'need to ask about rubric -- the only prices that should be considered are the opening price on first day and closing price on last day.
'rubric(cont) I dont think there is a reason to look at prices every day -- doing so would be unnecessary and would cause the script to take longer unnecessarly
'do we need to actually insert additional columns? or can we have data be inserted into existing blank columns?
'ask if we should always check if a variable is zero before trying to divide by a variable as a best practice 
End Sub
