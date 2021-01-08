Sub Stock_Analyst()

Dim Stock As String
Dim Year As Integer
Dim OpenPrice As Double
Dim ClosePrice As Double
Dim EntryRow As Integer
Dim Volume As Double
Dim ws As Worksheet

Dim BonusIncreaseStock As String
Dim BonusDecreaseStock As String
Dim BonusVolumeStock As String
Dim BonusIncrease As Double
Dim BonusDecrease As Double
Dim BonusVolume As Double

For Each ws In Worksheets
    'Set Header
    ws.Cells(1, 9) = "Ticker"
    ws.Cells(1, 10) = "Year"
    ws.Cells(1, 11) = "Yearly Change"
    ws.Cells(1, 12) = "Percent Change"
    ws.Cells(1, 13) = "Total Stock Volume"
    
    'Bonus - Greatest Increase, Decrease, Volume
    ws.Cells(1, 17) = "Ticker"
    ws.Cells(1, 18) = "Value"
    ws.Cells(2, 16) = "Greatest % Increase"
    ws.Cells(3, 16) = "Greatest % Decrease"
    ws.Cells(4, 16) = "Greatest Total Volume"
    'Fit column to width
    ws.Columns("P").AutoFit
    
    'Default variables
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    EntryRow = 2
    Stock = ws.Cells(2, 1).Value
    Year = Left(ws.Cells(2, 2).Value, 4)
    OpenPrice = ws.Cells(2, 3).Value
    Volume = ws.Cells(2, 7).Value
    BonusIncrease = 0
    BonusDecrease = 0
    BonusVolume = 0
    
    'Cells(2, 9) = Stock
    'Cells(2, 10) = Year

    
    For i = 2 To LastRow
        'Checking if Stock matches. If no, update stock and year, if yes, check year and update if different
        If ws.Cells(i, 1).Value <> Stock Then
            'Enter Stock, Year, Yearly Change, Percent Change, Total Volume
            ws.Cells(EntryRow, 9) = Stock
            ws.Cells(EntryRow, 10) = Year
            ws.Cells(EntryRow, 11) = ClosePrice - OpenPrice
            'Accounting for opening price of 0
            If OpenPrice <> 0 Then
                ws.Cells(EntryRow, 12) = (ClosePrice - OpenPrice) / OpenPrice
            Else
                ws.Cells(EntryRow, 12) = 0
            End If
            ws.Cells(EntryRow, 13) = Volume
            
            EntryRow = EntryRow + 1
           
            'Update Variables
            Stock = ws.Cells(i, 1).Value
            Year = Left(ws.Cells(i, 2).Value, 4)
            OpenPrice = ws.Cells(i, 3).Value
            Volume = ws.Cells(i, 7).Value
            
         Else
            If Left(ws.Cells(i, 2).Value, 4) <> Year Then
                'Enter Yearly, Percent Change, Total Volume
                ws.Cells(EntryRow, 9) = Stock
                ws.Cells(EntryRow, 10) = Year
                ws.Cells(EntryRow, 11) = ClosePrice - OpenPrice
                ws.Cells(EntryRow, 12) = (ClosePrice - OpenPrice) / OpenPrice
                ws.Cells(EntryRow, 13) = Volume
    
                
                EntryRow = EntryRow + 1
                
                Year = Left(ws.Cells(i, 2).Value, 4)
                
    
            Else
                Volume = Volume + ws.Cells(i, 7).Value
                ClosePrice = ws.Cells(i, 6).Value
                
            End If
        End If
        
        'Accounting for last row
        If i = LastRow Then
            ws.Cells(EntryRow, 9) = Stock
            ws.Cells(EntryRow, 10) = Year
            ws.Cells(EntryRow, 11) = ClosePrice - OpenPrice
            ws.Cells(EntryRow, 12) = (ClosePrice - OpenPrice) / OpenPrice
            ws.Cells(EntryRow, 13) = Volume
        End If
    
    Next i


        
    'Reformat Change columns
    LastRow = ws.Cells(Rows.Count, 11).End(xlUp).Row
    'MsgBox (LastRow)
    
    For x = 2 To LastRow
        ws.Cells(x, 11).FormatConditions.Delete            'Deleting existing conditional formatting
        'Red Cells
        ws.Cells(x, 11).FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, _
            Formula1:="=0"
        ws.Cells(x, 11).FormatConditions(1).Interior.Color = RGB(255, 0, 0)
        
        'Green Cells
        ws.Cells(x, 11).FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, _
            Formula1:="=0"
        ws.Cells(x, 11).FormatConditions(2).Interior.Color = RGB(0, 255, 0)
        
        'Format to 2 decimals and percent
        ws.Cells(x, 11).NumberFormat = "#,##0.00#"
        ws.Cells(x, 12).NumberFormat = "0.00%"
        
        
        'Bonus Greatest Increase
        If ws.Cells(x, 11).Value > BonusIncrease Then
            BonusIncrease = ws.Cells(x, 11).Value
            BonusIncreaseStock = ws.Cells(x, 9)
            
        End If
        
        'Bonus Greatest Decrease
        If ws.Cells(x, 11).Value < BonusDecrease Then
            BonusDecrease = ws.Cells(x, 11).Value
            BonusDecreaseStock = ws.Cells(x, 9)
        End If
        
        'Bonus Greatest Volume
        If ws.Cells(x, 13).Value > BonusVolume Then
            BonusVolume = ws.Cells(x, 13).Value
            BonusVolumeStock = ws.Cells(x, 9)
        End If
      
        
    Next x
    ws.Cells(2, 17) = BonusIncreaseStock
    ws.Cells(2, 18) = BonusIncrease
    ws.Cells(3, 17) = BonusDecreaseStock
    ws.Cells(3, 18) = BonusDecrease
    ws.Cells(4, 17) = BonusVolumeStock
    ws.Cells(4, 18) = BonusVolume
       
    
Next ws

End Sub


