# VBA-challenge
Completed by Graham Marsh

Option Explicit
Dim i As LongLong
Dim Ticker As String
Dim Summary_Table_Row As Integer
Dim ws As Worksheet
Dim Total_Stock_Volume As LongLong
Dim OpenPrice As Double
Dim ClosePrice As Double
Dim YearlyPriceChange As Double
Dim EndDate As Long
Dim StartDate As Long
Dim YearlyPercentageChange As Double
Dim maxValue As Double
Dim minValue As Double
Dim maxVolume As LongLong
Dim rng As Range
Dim cell As Range
Dim lastRow As Long

Sub RunAllMacros()

'A function to run all functions
    Call ConvertRangeToNumberOnWorksheetsWithLastRow
    Call CreateColumnHeadersOnAllSheets
    Call CreateTickerSummaryTable
    Call CreateTotalStockVolume
    Call CreatePriceChange
    Call FormatPriceChangeCells
    Call CreatePercentageChange
    Call FormatPercentageChange
    Call CreateGreatestSummaryTable
    Call FindMaxPercentValue
    Call FindMinPercentValue
    Call FindMaxVolumeValue
End Sub

Sub ConvertRangeToNumberOnWorksheetsWithLastRow()
  
    
    ' Loop through each worksheet in the workbook
    For Each ws In ThisWorkbook.Worksheets
        ' Determine the last row in column B
        lastRow = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row
        
        ' Set the range you want to convert to numbers
        Set rng = ws.Range("B1:B" & lastRow)
        
        ' Loop through each cell in the range
        For Each cell In rng
            ' Check if the cell value is numeric
            If IsNumeric(cell.Value) Then
                ' Convert the cell value to a number
                cell.Value = CDbl(cell.Value)
            End If
        Next cell
    Next ws
End Sub

Sub CreateColumnHeadersOnAllSheets()
      
    ' Loop through each worksheet
    For Each ws In ThisWorkbook.Worksheets
    
        ' Set the header row
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
    
    Next ws

End Sub

Sub CreateTickerSummaryTable()

    'Loop through each worksheet in the workbook
    For Each ws In ThisWorkbook.Worksheets

        'Define lastRow
        lastRow = Cells(Rows.Count, 1).End(xlUp).Row
        
        ' Keep track of the location for each ticker in the summary table
        Summary_Table_Row = 2

            'Loop through all tickers
            For i = 2 To lastRow

                'Check if we are still within the same ticker, if it is not then
                If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

                'Set the ticker
                Ticker = ws.Cells(i, 1).Value

                'Print the ticker in the Summary Table
                ws.Range("I" & Summary_Table_Row).Value = Ticker

                'Add one to the summary table row
                Summary_Table_Row = Summary_Table_Row + 1
                
                
                End If

            Next i

    Next ws

End Sub

Sub CreateTotalStockVolume()

'Loop through each worksheet in the workbook
For Each ws In ThisWorkbook.Worksheets

        'Define lastRow
        lastRow = Cells(Rows.Count, 1).End(xlUp).Row
        
        'Define start point for total stock volume
        Total_Stock_Volume = 0
        
        'Keep track of the location for each ticker in the summary table
        Summary_Table_Row = 2

            'Loop through all tickers
            For i = 2 To lastRow

                'Check if we are still within the same ticker, if it is not then
                If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

                    Total_Stock_Volume = Total_Stock_Volume + ws.Cells(i, 7).Value

                    'Print the ticker in the Summary Table
                    ws.Range("L" & Summary_Table_Row).Value = Total_Stock_Volume

                    'Add one to the summary table row
                    Summary_Table_Row = Summary_Table_Row + 1
                
                    Total_Stock_Volume = 0
                    
                Else
                
                Total_Stock_Volume = Total_Stock_Volume + ws.Cells(i, 7).Value
                
                
                End If

            Next i

    Next ws

End Sub

Sub CreatePriceChange()

 'Loop through each worksheet in the workbook
    For Each ws In ThisWorkbook.Worksheets
    
        'Define lastRow
        lastRow = Cells(Rows.Count, 1).End(xlUp).Row
        
        'Define EndDate
        EndDate = WorksheetFunction.Max(ws.Range("B2:B" & lastRow))
        
        'Define StartDate
        StartDate = WorksheetFunction.Min(ws.Range("B2:B" & lastRow))
        
        
        YearlyPriceChange = 0
        
        Summary_Table_Row = 2

            'Define loop
            For i = 2 To lastRow
    
                
                'Define condition of start date
                If ws.Cells(i, 2).Value = StartDate Then

                    OpenPrice = ws.Cells(i, 3).Value
            
                End If
        
                'Define condition of end date, calculate price change and enter in the summary table
                If ws.Cells(i, 2).Value = EndDate Then

                    ClosePrice = ws.Cells(i, 6).Value
        
                    YearlyPriceChange = (ClosePrice - OpenPrice)
            
                    ws.Range("J" & Summary_Table_Row).Value = YearlyPriceChange
            
                    Summary_Table_Row = Summary_Table_Row + 1
            
                    YearlyPriceChange = 0
            
                Else
        
                    YearlyPriceChange = (ClosePrice - OpenPrice)
            
                End If
        
            Next i
    
    Next ws
    
End Sub

Sub FormatPriceChangeCells()

    'Loop through each worksheet in the workbook
    For Each ws In ThisWorkbook.Worksheets
    
    'Define lastRow
    lastRow = Cells(Rows.Count, 9).End(xlUp).Row
        
        'Loop through each row
        For i = 2 To lastRow
        
            'Define the condition and color index
            If ws.Cells(i, 10).Value <= 0 Then
            
                ws.Cells(i, 10).Interior.ColorIndex = 3
                
            Else
            
                ws.Cells(i, 10).Interior.ColorIndex = 4
                
            End If
        
        Next i
        
    Next ws

End Sub

Sub CreatePercentageChange()
    
    'Loop through each worksheet in the workbook
    For Each ws In ThisWorkbook.Worksheets
    
        'Define lastRow
        lastRow = Cells(Rows.Count, 1).End(xlUp).Row
        
        'Define EndDate
        EndDate = WorksheetFunction.Max(ws.Range("B2:B" & lastRow))
        
        'Define StartDate
        StartDate = WorksheetFunction.Min(ws.Range("B2:B" & lastRow))
        
        
        YearlyPercentageChange = 0
        
        Summary_Table_Row = 2

            'Define loop
            For i = 2 To lastRow
    
                
                'Define condition of start date
                If ws.Cells(i, 2).Value = StartDate Then

                    OpenPrice = ws.Cells(i, 3).Value
            
                End If
        
                'Define condition of end date, calculate price change and enter in the summary table
                If ws.Cells(i, 2).Value = EndDate Then

                    ClosePrice = ws.Cells(i, 6).Value
        
                    YearlyPercentageChange = ((ClosePrice - OpenPrice) / OpenPrice)
            
                    ws.Range("K" & Summary_Table_Row).Value = YearlyPercentageChange
            
                    Summary_Table_Row = Summary_Table_Row + 1
            
                    YearlyPercentageChange = 0
            
                Else
        
                    YearlyPercentageChange = ((ClosePrice - OpenPrice) / OpenPrice)
            
                End If
        
            Next i
    
    Next ws
    
End Sub

Sub FormatPercentageChange()
    
    'Loop through each worksheet in the workbook
    For Each ws In ThisWorkbook.Worksheets
    
    'Define lastRow
    lastRow = Cells(Rows.Count, 9).End(xlUp).Row
    
        'Create loop to apply number format
        For i = 2 To lastRow
            
            'Define number format to apply
            ws.Cells(i, 11).NumberFormat = "0.00%"
            
        Next i
        
    Next ws
End Sub
Sub CreateGreatestSummaryTable()

    'Loop through each worksheet in the workbook
    For Each ws In ThisWorkbook.Worksheets
    
        'Define the headers for rows and columns of the Greatest Summary Table
        ws.Range("o2").Value = "Greatest % Increase"
        ws.Range("o3").Value = "Greatest % Decrease"
        ws.Range("o4").Value = "Greatest Total Volume"
        ws.Range("p1").Value = "Ticker"
        ws.Range("q1").Value = "Value"
        
    Next ws
    
End Sub
Sub FindMaxPercentValue()

'Loop through each worksheet in the workbook
For Each ws In ThisWorkbook.Worksheets
    
    'Define lastRow
    lastRow = Cells(Rows.Count, 9).End(xlUp).Row
     
    'Define maximum percent value
    maxValue = Application.WorksheetFunction.Max(ws.Range("K2:K" & lastRow))

        'Create loop to find maximum value
        For i = 2 To lastRow
       
            'Define condition to find maximum value
            If ws.Cells(i, 11).Value = maxValue Then

                'Define the cell to insert the maximum percent value
                ws.Range("Q2").Value = ws.Cells(i, 11).Value
                
                'Define format for maximum percent value
                ws.Range("Q2").NumberFormat = "0.00%"
                
                'Define the cell to insert the ticker of the maximum percent value
                ws.Range("P2").Value = ws.Cells(i, 9).Value
            
            End If
            
        Next i
         
    Next ws
            
End Sub

Sub FindMinPercentValue()

'Loop through each worksheet in the workbook
For Each ws In ThisWorkbook.Worksheets
    
    'Define lastRow
    lastRow = Cells(Rows.Count, 9).End(xlUp).Row
    
    'Define minimum percent value
    minValue = Application.WorksheetFunction.Min(ws.Range("K2:K" & lastRow))
    
        'Create loop to find minimum value
        For i = 2 To lastRow
            
            'Define condition to find minimum value
            If ws.Cells(i, 11).Value = minValue Then

                'Define the cell to insert the minimum percent value
                ws.Range("Q3").Value = ws.Cells(i, 11).Value
                
                'Define format for maximum percent value
                ws.Range("Q3").NumberFormat = "0.00%"
                
                'Define the cell to insert the ticker of the minimum percent value
                ws.Range("P3").Value = ws.Cells(i, 9).Value
            
            End If
            
        Next i
         
    Next ws
            
End Sub

Sub FindMaxVolumeValue()

'Loop through each worksheet in the workbook
For Each ws In ThisWorkbook.Worksheets

    'Define lastRow
    lastRow = Cells(Rows.Count, 9).End(xlUp).Row
    
    'Define maximum volume
    maxVolume = Application.WorksheetFunction.Max(ws.Range("L2:L" & lastRow))

        'Create loop to find maximum volume value
        For i = 2 To lastRow
            
            'Define condition to find maximum volume value
            If ws.Cells(i, 12).Value = maxVolume Then

                'Define the cell to insert the maximum volume value
                ws.Range("Q4").Value = ws.Cells(i, 12).Value
                
                'Define the cell to insert the ticker of the maximum volume value
                ws.Range("P4").Value = ws.Cells(i, 9).Value
            
            End If
            
        Next i
         
    Next ws
            
End Sub








