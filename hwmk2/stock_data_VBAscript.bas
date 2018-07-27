Attribute VB_Name = "Module1"
Sub StockAnalysis()

' Making a variable for worksheet
Dim page As Worksheet

' Looping through each worksheet
For Each page In ActiveWorkbook.Worksheets
    page.Activate

' Setting headers for displaying values
    Cells(1, 9).Value = "Ticker"
    Cells(1, 10).Value = "Yearly Change"
    Cells(1, 11).Value = "Percent Change"
    Cells(1, 12).Value = "Total Stock Volume"
    
    Cells(1, 16).Value = "Ticker"
    Cells(1, 17).Value = "Value"
    Cells(2, 15).Value = "Greatest % Increase"
    Cells(3, 15).Value = "Greatest % Decrease"
    Cells(4, 15).Value = "Greatest Total Volume"

' Declaring variables needed
    Dim volume As Double
    Dim yearly_change As Single
    Dim percent_change As Double
    Dim ticker As String
    Dim open_price As Double
    Dim close_price As Double
     
' For new rows on summary table
    Dim row_count As Double
    
' Setting intial values to variables
    volume = 0
    row_count = 2

' Setting opening price
    open_price = Cells(2, 3).Value
    
' Determining last row
    last_row = page.Cells(Rows.Count, 1).End(xlUp).Row
    
    ' Looping through cells of each sheet
        For i = 2 To last_row
        
        ' If cell not equal to previous
            If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
                ' Setting ticker name based on previous cell and begins second list 8 columns away
                    ticker = Cells(i, 1).Value
                    Cells(row_count, 9).Value = ticker
                ' Setting close price
                    close_price = Cells(i, 6).Value
                ' Determing and printing yearly change with correct format
                    yearly_change = close_price - open_price
                    Cells(row_count, 10).Value = yearly_change
                    ' Determing percent change and dealing with dividing by 0
                        If (open_price = 0 And close_price = 0) Then
                            percent_change = 0
                        ElseIf (open_price = 0 And close_price <> 0) Then
                            percent_change = 1
                        Else
                            percent_change = yearly_change / open_price
                        ' Printing value with correct format
                            Cells(row_count, 11).Value = percent_change
                            Cells(row_count, 11).NumberFormat = "0.00%"
                        End If
                ' Determing and printing total volume
                    volume = volume + Cells(i, 7)
                    Cells(row_count, 12).Value = volume
                ' Creating next row on summary table
                    row_count = row_count + 1
                ' Resetting open price and volume for next ticker
                    open_price = Cells(i + 1, 3)
                    volume = 0
            Else
                    volume = volume + Cells(i, 7).Value
            End If
            
        Next i
          
' Determine the last row of yearly change column per sheet
    last_yearly_change = page.Cells(Rows.Count, 9).End(xlUp).Row
    ' Setting the colors of the cells for yearly change
        For j = 2 To last_yearly_change
        
            If (Cells(j, 10).Value > 0 Or Cells(j, 10).Value = 0) Then
                Cells(j, 10).Interior.ColorIndex = 4
            ElseIf Cells(j, 10).Value < 0 Then
                Cells(j, 10).Interior.ColorIndex = 3
            End If
            
        Next j
    ' Determining minimum and maximum values
        For k = 2 To last_yearly_change
        
            If Cells(k, 11).Value = Application.WorksheetFunction.Max(page.Range("K2:K" & last_yearly_change)) Then
                Cells(2, 16).Value = Cells(k, 9).Value
                Cells(2, 17).Value = Cells(k, 11).Value
            ElseIf Cells(k, 11).Value = Application.WorksheetFunction.Min(page.Range("K2:K" & last_yearly_change)) Then
                Cells(3, 16).Value = Cells(k, 9).Value
                Cells(3, 17).Value = Cells(k, 11).Value
            ElseIf Cells(k, 12).Value = Application.WorksheetFunction.Max(page.Range("L2:L" & last_yearly_change)) Then
                Cells(4, 16).Value = Cells(k, 9).Value
                Cells(4, 17).Value = Cells(k, 12).Value
            End If
        
        Next k
    
    ' Formats new columns to content width
    Columns("I:Q").Select
    Selection.EntireColumn.AutoFit
    
    Next page

End Sub
