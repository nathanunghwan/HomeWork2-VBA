Attribute VB_Name = "Module1"

Sub Stock_Analysis()

 Dim Count As Integer
 Dim i As Long
 Dim j As Integer
 Dim lastRow As Long
 Dim ticker_num As Integer
 
 Count_sheet = ActiveWorkbook.Worksheets.Count
         
 'Loop to operate on all Sheet with one macro button(Open from sheet1 to sheet3 in order)
 For j = 1 To Count_sheet
    'Active Worksheet
     Worksheets(j).Activate
    
     'Set result Index
     Range("I1").Value = "Ticker"
     Range("J1").Value = "Yearly Change"
     Range("K1").Value = "Percent Change"
     Range("L1").Value = "Total Stock Volume"
     Range("Q1").Value = "Ticker"
     Range("R1").Value = "Value"
     Range("P2").Value = "Greatest % Increase"
     Range("P3").Value = "Greatest % Decrease"
     Range("P4").Value = "Greatest Total Volume"
    
     'Find the last row number where the raw data was entered
     lastRow = Cells(Rows.Count, 1).End(xlUp).Row
     
     'Set collection data initial value
     Open_Price = Cells(2, 3).Value                   'The opening price at the beginning of the year
     ticker_num = 2                                          'Assign a row number for recording result data
    
    'Loop to last row of raw data entered
     For i = 2 To lastRow
         'Accumulation of stock volume
         Stock_Volume = Stock_Volume + Cells(i, 7)
         
         'Search for condition to see if the ticker in the raw data changes
         If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
         
             Close_Price = Cells(i, 6).Value                                                  'assign close price
             Cells(ticker_num, 9).Value = Cells(i, 1)                                     'write result ticker data
             Cells(ticker_num, 10).Value = Close_Price - Open_Price           'write result Yearly change
            
            'Set Yearly Change interior color
             If Cells(ticker_num, 10).Value > 0 Then
                 Cells(ticker_num, 10).Interior.ColorIndex = "4"
                 
             ElseIf Cells(ticker_num, 10).Value < 0 Then
                 Cells(ticker_num, 10).Interior.ColorIndex = "3"
                
             Else                                                                                       'if open price = close price
                 Cells(ticker_num, 10).Interior.ColorIndex = "0"
                
             End If
             
             'Write percent change, total stock volume
             Cells(ticker_num, 11).Value = (Close_Price - Open_Price) / Open_Price
             Cells(ticker_num, 12).Value = Stock_Volume
             
             'reseting
             ticker_num = ticker_num + 1                                               '+1 to increase row number for writing result
             Open_Price = Cells(i + 1, 3).Value                                        'setting i+1 open_price
             Stock_Volume = 0                                                                'reset stock volume for accumualting to next ticker
             
             'set number format for looking easely
             Cells(ticker_num, 10).NumberFormat = "[$$]#,##0.00 "
             Cells(ticker_num, 11).NumberFormat = "0.00%;[red]-0.00%"
             Cells(ticker_num, 12).NumberFormat = "#,##0"
         
         End If
         
     Next i
    
    'Funtionality(Max,Min) to script to return the value of the stock with the "Greatest % increase","Greatest%decrease", and "Greatest total volum
     Range("R2").Value = WorksheetFunction.Max(Range("K2:" & "K" & lastRow))
     Range("R3").Value = WorksheetFunction.Min(Range("K2:" & "K" & lastRow))
     Range("R4").Value = WorksheetFunction.Max(Range("L2:" & "L" & lastRow))
     
     'Funtionality(Index_Match) to script to return the ticker of the stock with the "Greatest % increase","Greatest%decrease", and "Greatest total volum
     Range("Q2").Value = WorksheetFunction.Index(Range("I2:" & "I" & lastRow), WorksheetFunction.Match(Range("R2").Value, Range("K2:" & "K" & lastRow), 0))
     Range("Q3").Value = WorksheetFunction.Index(Range("I2:" & "I" & lastRow), WorksheetFunction.Match(Range("R3").Value, Range("K2:" & "K" & lastRow), 0))
     Range("Q4").Value = WorksheetFunction.Index(Range("I2:" & "I" & lastRow), WorksheetFunction.Match(Range("R4").Value, Range("L2:" & "L" & lastRow), 0))
     
     'set number format for looking easely
     Range("R2").NumberFormat = "0.00%;[red]-0.00%"
     Range("R3").NumberFormat = "0.00%;[red]-0.00%"
     Range("R4").NumberFormat = "#,##0"
     
     'set interior color for looking result easely
     Range("I1:L1, P1:R1, P3:R3").Select
     Selection.Interior.ColorIndex = "37"
     
     'set colum wideth for looking result easely
     Columns("I:R").EntireColumn.AutoFit
     Range("I:I,Q:Q").Select
     Selection.ColumnWidth = 11
 

 Next j
 
 'Activate the first worksheet after completing the last worksheet task
 Worksheets(1).Activate
 Range("A1").Select
 
End Sub

