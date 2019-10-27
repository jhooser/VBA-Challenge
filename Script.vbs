Sub StockingUp()
'Establish variables that will be compared across the whole sheet
greatestper = 0
worstper = 0
greatestvol = 0

'For every worksheet in the excel files
For Each ws In Worksheets
    'label the columns correctly
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"
    Dim yearper, yeardif, yearopen As Long
    
    'Set a starting point for our row calculator
    Row_count = 2
    'Create the last row counter
    Dim LastRow As Long
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    'Establish the first last row value
    yearopen = ws.Cells(2, 3).Value
    'Set the ticker counter
    tickcount = 0
    For i = 2 To LastRow

    'If the cells before the next cell doesn't match up:
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            'Increase the tick counter
            tickcount = tickcount + ws.Cells(i, 7).Value
            'brand name is displayed
            ws.Cells(Row_count, 9).Value = ws.Cells(i, 1)
            
            'pull your year close value from the table
            yearclose = ws.Cells(i, 6).Value
            
            'display year difference between year close and year open
            yeardif = Round(yearclose - yearopen, 2)
            ws.Cells(Row_count, 10).Value = yeardif
            'Calculate the percent change and display it
            If yearopen = 0 Then
                yearper = 0
            Else
                yearper = Round((yeardif / yearopen) * 100, 2)
            End If
            ws.Cells(Row_count, 11).Value = yearper
            'mark the cells as green if over 0
            If ws.Cells(Row_count, 11).Value > 0 Then
                ws.Cells(Row_count, 11).Interior.ColorIndex = 4
            'otherwise mark them as red
            Else
                ws.Cells(Row_count, 11).Interior.ColorIndex = 3
            End If
            'Compare year percentage to greatest year percentage and save the higher number
            If yearper > greatestper Then
                greatestper = yearper
                greatestpern = ws.Cells(i, 1)
            End If
            'Compare year percentage to worst year percentage and save the lower number
            If yearper < worstper Then
                worstper = yearper
                worstpern = ws.Cells(i, 1)
            End If
            'Place the total number of volume traded in the correct column
            ws.Cells(Row_count, 12).Value = tickcount
            'Compare tickcount to highest traded value and save the higher number
            If tickcount > greatestvol Then
                greatestvol = tickcount
                greatestvoln = ws.Cells(i, 1)
            End If
            tickcount = 0
            'Row count number is increased
            Row_count = Row_count + 1
            'Creat new year open value
            yearopen = ws.Cells(i + 1, 3)
        Else
            'if it is the same ticker, just need to increase the ticker counter
            tickcount = tickcount + ws.Cells(i, 7).Value
        End If
    Next i
   ws.Columns("I:L").AutoFit
Next ws

'Print the greatest percentage, worst percentage, and greatest volume numbers in the first sheet.
Cells(2, 15).Value = "Greatest Percentage"
Cells(3, 15).Value = "Worst Percentage"
Cells(4, 15).Value = "Greatest Volume"
Cells(1, 16).Value = "Ticker"
Cells(1, 17).Value = "Value"
Cells(2, 16).Value = greatestpern
Cells(2, 17).Value = greatestper
Cells(3, 16).Value = worstpern
Cells(3, 17).Value = worstper
Cells(4, 16).Value = greatestvoln
Cells(4, 17).Value = greatestvol
Columns("O:Q").AutoFit
End Sub

Sub Reset()
For Each ws In Worksheets
  ws.Columns("I:Q").EntireColumn.Delete
Next ws
End Sub