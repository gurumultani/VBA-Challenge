Attribute VB_Name = "Module1"
Sub stockanalysis():
'Declairing variables
    Dim ws As Worksheet
    Dim lastrow As Long
    Dim ticker As String
    Dim openingprice As Double
    Dim closingprice As Double
    Dim yearlychange As Double
    Dim percentchange As Double
    Dim totalvolume As Double
    Dim summarytable As Range
    Dim row As Long
    Dim increase_number As Double
    Dim decrease_number As Double
    Dim volume_number As Double


  'loop runs through all sheets simultaneously
  For Each ws In ThisWorkbook.Worksheets
  
  
'find the last row of data in column
 lastrow = ws.Cells(ws.Rows.Count, 1).End(xlUp).row
 
 'Set up the summary table headers
  ws.Cells(1, 9).Value = "Ticker"
  ws.Cells(1, 10).Value = "Yearly Change"
  ws.Cells(1, 11).Value = "Percent Change"
  ws.Cells(1, 12).Value = "Total Stock Volume"
  ws.Range("O2").Value = "Greatest % Increase"
  ws.Range("O3").Value = "Greatest % Decrease"
  ws.Range("O4").Value = "Greatest Total Volume"
  ws.Cells(1, 16).Value = "Ticker"
  ws.Cells(1, 17).Value = "Value"
  
  'Set the summary table range
  
  Set summarytable = ws.Range("I2:L" & lastrow)
  
  'Initialise Variables
  ticker = ws.Cells(2, 1).Value
  openingprice = ws.Cells(2, 3).Value
  totalvolume = 0
  row = 1
  
   
  'Loop through all rows of data
    For i = 1 To lastrow
    
    
      
     'Check if the ticker symbol has changed
     If ws.Cells(i + 1, 1).Value <> ticker Then
     
          'Get the closing price
           closingprice = ws.Cells(i, 6).Value
         
         'Calculate the yearly change and percent change
          yearlychange = closingprice - openingprice
          percentchange = (yearlychange / openingprice) * 100
          
          'Output the results in summary table
          summarytable.Cells(row, 1).Value = ticker
          summarytable.Cells(row, 2).Value = yearlychange
          summarytable.Cells(row, 3).Value = percentchange & "%"
          summarytable.Cells(row, 4).Value = totalvolume
          
          'Reset variables for the next ticker symbol
           row = row + 1
           ticker = ws.Cells(i + 1, 1).Value
           openingprice = ws.Cells(i + 1, 3).Value
           totalvolume = 0
        End If
        
        'Accumulate the total stock volume
        totalvolume = totalvolume + ws.Cells(i + 1, 7).Value
        
    Next i
    
        'Applying conditional formatting to highlight positive yearly change
         summarytable.Columns("B").FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, Formula1:="0"
         summarytable.Columns("B").FormatConditions(1).Interior.ColorIndex = 4 ' Green color for positive change
         
         'Apply Conditional formatting to highlight negative yearly change
         summarytable.Columns("B").FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, Formula1:="0"
         summarytable.Columns("B").FormatConditions(2).Interior.ColorIndex = 3 ' Red color for negative change
         
         'Applying conditional formatting to highlight positive percent change
         summarytable.Columns("C").FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, Formula1:="0"
         summarytable.Columns("C").FormatConditions(1).Interior.ColorIndex = 4 ' Green color for positive change
         
         'Apply Conditional formatting to highlight negative percent change
         summarytable.Columns("C").FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, Formula1:="0"
         summarytable.Columns("C").FormatConditions(2).Interior.ColorIndex = 3 ' Red color for negative change
         
  
         ' take the max and min and place them in a separate part in the worksheet
        ws.Range("Q2") = "%" & WorksheetFunction.Max(ws.Range("K2:K" & lastrow)) * 100
        ws.Range("Q3") = "%" & WorksheetFunction.Min(ws.Range("K2:K" & lastrow)) * 100
        ws.Range("Q4") = WorksheetFunction.Max(ws.Range("L2:L" & lastrow))
         
        ' returns one less because header row not a factor
        increase_number = WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("K2:K" & lastrow)), ws.Range("K2:K" & lastrow), 0)
        decrease_number = WorksheetFunction.Match(WorksheetFunction.Min(ws.Range("K2:K" & lastrow)), ws.Range("K2:K" & lastrow), 0)
        volume_number = WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("L2:L" & lastrow)), ws.Range("L2:L" & lastrow), 0)

      ' final ticker symbol for  total, greatest % of increase and decrease, and average
        ws.Range("P2") = ws.Cells(increase_number + 1, 9)
        ws.Range("P3") = ws.Cells(decrease_number + 1, 9)
        ws.Range("P4") = ws.Cells(volume_number + 1, 9)
Next ws

    End Sub

