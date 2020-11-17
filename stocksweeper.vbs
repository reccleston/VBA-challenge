' Ryan Eccleston-Murdock
' VBA HW - HW3 
' 21 November 

' Purpose: Get and report ticker, yearly change, percent change, and total volume of a given stock of a given year as well as an
' extrema summary 

sub stocksweeper()

    for each sheet in Worksheets
        call make_summary_table(sheet)
        call stock_extrema(sheet)
    next sheet

end sub

sub stock_extrema(sheet)

    last_row = sheet.cells(rows.count, "k").end(xlUp).row - 1

    ' Labeling/ Formatting of extrema
    sheet.range("n2").value = "Greatest % Increase"
    sheet.range("n3").value = "Greatest % Decrease"
    sheet.range("n4").value = "Greatest Volume"
    sheet.range("o1").value = "<ticker>"
    sheet.range("p1").value = "Value"
    sheet.range("p2:p3").numberformat = "0.00%"

    ' Value 
    greatest_percent_increase = worksheetfunction.max(sheet.range("k:k"))
    greatest_percent_decrease = worksheetfunction.min(sheet.range("k:k"))
    greatest_vol = worksheetfunction.max(sheet.range("l:l"))

    sheet.range("p2").value = greatest_percent_increase
    sheet.range("p3").value = greatest_percent_decrease
    sheet.range("p4").value = greatest_vol

    ' Ticker Name
    increase_ticker = application.index(sheet.range("i2:i" & last_row), application.match(sheet.range("p2").value, sheet.range("k2:k" & last_row), 0))
    decrease_ticker = application.index(sheet.range("i2:i" & last_row), application.match(sheet.range("p3").value, sheet.range("k2:k" & last_row), 0))
    vol_ticker = application.index(sheet.range("i2:i" & last_row), application.match(sheet.range("p4").value, sheet.range("l2:l" & last_row), 0))

    sheet.range("o2").value = increase_ticker
    sheet.range("o3").value = decrease_ticker 
    sheet.range("o4").value = vol_ticker

end sub

sub make_summary_table(sheet)

    dim percent_change as double

    sheet.range("i1").value = "<ticker>"
    sheet.range("j1").value = "<yearly change>"
    sheet.range("k1").value = "<percent change>"
    sheet.range("k:k").numberformat = "0.00%"
    sheet.range("l1").value = "<total volume>"

    ' Populating summary table 
    last_row = sheet.cells(rows.count, "a").end(xlUp).row
    summary_table_row = 2
    stock_volume = 0 
    ticker_type = 0 

    for row = 2 to last_row
        if sheet.cells(row + 1, 1).value <> sheet.cells(row, 1).value then
            ' Get and record ticker name
            ticker = sheet.cells(row, 1).value 
            sheet.range("i" & summary_table_row).value = ticker

            ' Get and record change over time 
            open_price = sheet.cells(row - ticker_type, 3).value
            close_price = sheet.cells(row, 6).value
            change = close_price - open_price
            sheet.range("j" & summary_table_row).value = change

            ' Get and record % increase
            if open_price = 0 then ' failsafe against opening price starting at 0
                open_price = 1
            end if 
            percent_change = change / open_price
            sheet.range("k" & summary_table_row).value = percent_change

            ' Conditional color formatting
            if percent_change < 0 then 
                sheet.range("j" & summary_table_row).interior.colorindex = 3
            else 
                sheet.range("j" & summary_table_row).interior.colorindex = 4
            end if

            'Increase summary table row 
            summary_table_row = summary_table_row + 1 
            ticker_type = 0
            stock_volume = 0
        else 
            ' Tally number of similar tickers
            ticker_type = ticker_type + 1 

            ' Calculate sum of total volume
            stock_volume = stock_volume + sheet.cells(row, 7).value 
        end if

        ' Populate total stock volume cell  
        sheet.range("l" & summary_table_row) = stock_volume
    next row 

end sub