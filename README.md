# challenge-2 vba code








Sub StockDataAnalysis()
    ' Set initial variables
    Dim ticker As String
    Dim yearly_change As Double
    Dim percent_change As Double
    Dim total_stock_volume As Double
    total_stock_volume = 0

    ' Set row counter for summary table
    Dim summary_table_row As Integer
    summary_table_row = 2

    ' Set initial open price and check if it's a valid numeric value
    Dim open_price As Double
    If IsNumeric(Cells(2, 3).Value) Then
        open_price = Cells(2, 3).Value
    Else
        MsgBox "Invalid opening price in cell (2, 3). Please check the data."
        Exit Sub
    End If

    ' Get the last row number of the data
    Dim last_row As Long
    last_row = Cells(Rows.Count, 1).End(xlUp).Row

    ' Variables to track greatest increase, decrease, and volume
    Dim greatest_increase As Double
    Dim greatest_decrease As Double
    Dim greatest_volume As Double
    Dim greatest_increase_ticker As String
    Dim greatest_decrease_ticker As String
    Dim greatest_volume_ticker As String

    ' Initialize variables with initial values
    greatest_increase = 0
    greatest_decrease = 0
    greatest_volume = 0
    greatest_increase_ticker = ""
    greatest_decrease_ticker = ""
    greatest_volume_ticker = ""

    ' Loop through all stock data
    For i = 2 To last_row

        ' Check if we are still within the same stock ticker
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

            ' Set ticker symbol
            ticker = Cells(i, 1).Value

            ' Calculate yearly change and percent change
            yearly_change = Cells(i, 6) - open_price
            If open_price <> 0 Then
                percent_change = yearly_change / open_price
            Else
                percent_change = 0
            End If

            ' Add to total stock volume
            total_stock_volume = total_stock_volume + Cells(i, 7).Value

            ' Print results in summary table
            Range("I" & summary_table_row).Value = ticker
            Range("J" & summary_table_row).Value = yearly_change
            Range("K" & summary_table_row).Value = percent_change
            Range("L" & summary_table_row).Value = total_stock_volume

            ' Apply conditional formatting to change font color
            Range("K" & summary_table_row).NumberFormat = "0.00%"
            If yearly_change < 0 Then
                Range("J" & summary_table_row).Font.Color = RGB(255, 0, 0) ' Red color for negative change
            Else
                Range("J" & summary_table_row).Font.Color = RGB(0, 176, 80) ' Green color for positive change
            End If

            ' Check for greatest increase, decrease, and volume
            If percent_change > greatest_increase Then
                greatest_increase = percent_change
                greatest_increase_ticker = ticker
            End If

            If percent_change < greatest_decrease Then
                greatest_decrease = percent_change
                greatest_decrease_ticker = ticker
            End If

            If total_stock_volume > greatest_volume Then
                greatest_volume = total_stock_volume
                greatest_volume_ticker = ticker
            End If

            ' Add one to the summary table row counter
            summary_table_row = summary_table_row + 1

            ' Reset total stock volume and set new open price for the next stock ticker
            total_stock_volume = 0
            If IsNumeric(Cells(i + 1, 3).Value) Then
                open_price = Cells(i + 1, 3).Value
            Else
                MsgBox "Invalid opening price in cell (" & i + 1 & ", 3). Please check the data."
                Exit Sub
            End If

        Else

            ' Add to total stock volume if still within the same stock ticker
            total_stock_volume = total_stock_volume + Cells(i, 7).Value

        End If

    Next i

    ' Print greatest increase, decrease, and volume
    Range("N2").Value = "Greatest % Increase"
    Range("N3").Value = "Greatest % Decrease"
    Range("N4").Value = "Greatest Total Volume"
    Range("O2").Value = greatest_increase_ticker
    Range("O3").Value = greatest_decrease_ticker
    Range("O4").Value = greatest_volume_ticker
    Range("P2").Value = greatest_increase
    Range("P3").Value = greatest_decrease
    Range("P4").Value = greatest_volume
    Range("P2:P3").NumberFormat = "0.00%"
    Range("P2:P4").HorizontalAlignment = xlRight

    ' Apply conditional formatting to change font color of greatest increase and decrease
    Range("P2").Font.Color = RGB(0, 176, 80) ' Green color for greatest increase
    Range("P3").Font.Color = RGB(255, 0, 0) ' Red color for greatest decrease

    MsgBox "Stock data analysis completed."

End Sub
