Sub BotClickTest2()
'
' BotClickTest Macro
'
Dim cl As Long, rw As Long, source As String, x As Range
   Dim i As Long, n As Long
   Dim extwbk As Workbook, twb As Workbook
   Application.ScreenUpdating = False
   Application.DisplayAlerts = False
   Application.EnableEvents = False
   


'Code here to split out sends vs clicks and add sheets
    Sheets.Add After:=ActiveSheet '<add sheets 2
    Sheets.Add After:=ActiveSheet '<add sheets 3
Sheets(1).Activate
  n = Cells(Rows.Count, "A").End(xlUp).row '<count rows

Range("A1:G" & n).AutoFilter Field:=1, Criteria1:= _
        "sent_campaign"
    Range("A1:G" & n).SpecialCells(xlCellTypeVisible).Copy
    Sheets(2).Paste
    
Range("A1:G" & n).AutoFilter Field:=1, Criteria1:= _
        "click"
    Range("A1:G" & n).SpecialCells(xlCellTypeVisible).Copy
    Sheets(3).Paste


Sheets(2).Name = "Sent" '<ensures second sheet is called sent
Sheets(3).Name = "Click" '<ensures third sheet is called click


Sheets("Sent").Activate
    Columns("B:B").Select
    Selection.Cut
    Columns("H:H").Select
    Selection.Insert Shift:=xlToRight 'pastes datetime of send to end of sheet
    
    
'<vlookup to pull in send datetime formula here:
Sheets("Click").Activate
Cells(1, 8).Value = "Sent datetime" '<add header
Cells(2, 8).FormulaR1C1 = "=VLOOKUP(RC[-3],Sent!C[-4]:C[-1],4,0)" '<vlookup recordID against sheet1 send datetimes
  n = Cells(Rows.Count, "A").End(xlUp).row '<count rows on sheet3
               Range("H2").Select
              Range("H2").AutoFill Destination:=Range("H2:H" & n) '<fill formula down to end
          Columns("H").Copy
          Columns("H").PasteSpecial xlPasteValues

 
'Calculate time difference between send and click datetime
    Cells(1, 9).Value = "Bot Click Test" '<adding header
    Cells(2, 9).FormulaR1C1 = "=IF(RC[-7]<RC[-1]+0.001388,""Bot"",""True Click"")" '<if the click is within 2 mins of send time mark as bot 0.001388 is 2 mins in excel datetime format
              Range("I2").Select
              Range("I2").AutoFill Destination:=Range("I2:I" & n) '<fill formula down to end
          Columns("I").Copy
          Columns("I").PasteSpecial xlPasteValues
 
    
    With Application
        .ScreenUpdating = True
        .EnableEvents = True
        .DisplayAlerts = True
        .Calculation = xlCalculationAutomatic
    End With
    
    
End Sub


