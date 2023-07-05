# VBA-Challenge
challengeSubmission
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

        - all the above codes for the bonus were provided by One on One tutor named:  Imaad Fakier


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


        - above code was the output from google, stack overflow which means ... continuous searching online


        - Also, most of the code and some commands are directly from the activities which were done in the regular lecture; like credit_card activity, census activity from where I directly taken the code for 'loop runs through all sheets simultaneously'
