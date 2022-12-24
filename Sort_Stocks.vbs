'****************************************
'*  A VB Script to calulate stock data  *
'*       -Justine Pile- [Dec 2022]      *
'****************************************

Sub StockData()
    
    'Get total number of rows of stock data
    No_of_Rows = Range("A2").End(xlDown).Row
    
    'Set j = 1 so that data output is offset by 1 row
    j = 1
    'Set columns headers
    Range("I1").Value = "Ticker"
    Range("J1").Value = "Total Stock Volume"
    Range("K1").Value = "Yearly Change"
    Range("L1").Value = "Percent Change"
    Range("O1").Value = "Ticker"
    Range("P1").Value = "Value"
    Range("N2").Value = "Greatest % Increase"
    Range("N3").Value = "Greatest % Decrease"
    Range("N4").Value = "Greatest Total Volume"
    
    
    'Loop through all rows of data
    For i = 2 To No_of_Rows
    
        If Year_Open = 0 Then
            Year_Open = Cells(i, 3)
        End If
        
        'If cell below is same as current cell then it is
        'the same stock so we will add volume of the stock
        If Cells(i, 1) = Cells(i + 1, 1) Then
            vol = Cells(i, 7) + vol
            Ticker = Cells(i, 1)
            
        'Otherwise cell below is different from the current cell
        '(i.e. last cell of same type)
        Else
            'Add the final row of that stock to amt
            vol = Cells(i, 7) + vol
            
            'Increment j by one (1) so data can be output to a new line
            j = j + 1
            
            'Get value of last closing price and calculate diff from initial opening
            Year_Close = Cells(i, 6)
            Year_Change = Year_Close - Year_Open
            
            'Output ticker, total volume, and yearly change to worksheet
            Cells(j, 9).Value = Ticker
            Cells(j, 10).Value = vol
            Cells(j, 11).Value = Year_Change
            
            'Calculate the yearly change percentage
            Cells(j, 12).Value = FormatPercent(Year_Change / Year_Open)
            
            'Reset value of amt and Year_Open
            vol = 0
            Year_Open = 0
            
        End If
        
    Next i
    
    'Find the greastest increase and decrease and output to sheet
    Set RngPct = Range("L2:L" & Rows.Count)
    Greatest_Inc = Application.WorksheetFunction.Max(RngPct)
    Greatest_Dec = Application.WorksheetFunction.Min(RngPct)
    Range("P2").Value = FormatPercent(Greatest_Inc)
    Range("P3").Value = FormatPercent(Greatest_Dec)
    
    'Find the greastest volume output to sheet
    Set RngVol = Range("J2:L" & Rows.Count)
    Greatest_Vol = Application.WorksheetFunction.Max(RngVol)
    Range("P4") = Greatest_Vol
    
    'Find associated ticker symbols for greatest increase/decrease
    Macro_Rows = Range("L2").End(xlDown).Row
    For k = 2 To Macro_Rows
        If Cells(k, 12).Value = Range("P2").Value Then
            Increase_Ticker = Cells(k, 9).Value
        End If
        If Cells(k, 12).Value = Range("P3").Value Then
            Decrease_Ticker = Cells(k, 9).Value
        End If
    Next k
    
    'Find associated ticker symbols for greatest volum
    For m = 2 To Macro_Rows
        If Cells(m, 10).Value = Range("P4").Value Then
            Greatest_Vol = Cells(m, 9).Value
        End If
    Next m
    
    'Output ticker symbols for greatest increase/decrease and vol
    Range("O2").Value = Increase_Ticker
    Range("O3").Value = Decrease_Ticker
    Range("O4").Value = Greatest_Vol
    
    'Format volumes as numbers with comma seperators
    Columns("J:J").Select
    Selection.NumberFormat = "#,##0"
    Range("P4").Select
    Selection.NumberFormat = "#,##0"
    
    'Reposition/reset the sheet position and cursor selection
    Range("I1").Select
    
    'Conditional formatting for positive and negative changes
    For n = 2 To Macro_Rows
        If Cells(n, 11).Value > 0 Then
            Cells(n, 11).Interior.ColorIndex = 4
        ElseIf Cells(n, 11).Value < 0 Then
            Cells(n, 11).Interior.ColorIndex = 3
        End If
    Next n
    
    'Autofit all columns
    Cells.EntireColumn.AutoFit
        
End Sub
