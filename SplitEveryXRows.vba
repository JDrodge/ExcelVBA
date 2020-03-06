Sub SplitRows()
'Definitions
Dim inputFile As String, inputWb As Workbook
    Dim lastRow As Long, row As Long, n As Long
    Dim newCSV As Workbook

'Find last row in column A, change as appropriate
With ActiveWorkbook.Worksheets(1)
    lastRow = .Cells(Rows.Count, "A").End(xlUp).row

    Set newCSV = Workbooks.Add

    'Copy headers and every 7000 rows stepped into a new file
    n = 0
    For row = 2 To lastRow Step 7000
        n = n + 1
        .Rows(1).EntireRow.Copy newCSV.Worksheets(1).Range("A1")
        .Rows(row & ":" & row + 7000 - 1).EntireRow.Copy newCSV.Worksheets(1).Range("A2")

        'Save in same folder as input workbook with .xlsx replaced by (n).csv
        newCSV.SaveAs fileName:=n & ".csv", FileFormat:=xlCSV, CreateBackup:=False
    Next
End With

newCSV.Close saveChanges:=False

End Sub
