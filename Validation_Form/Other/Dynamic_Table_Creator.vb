Sub DynamicRange()
'Best used when you want to include all data stored on the spreadsheet

Dim sht As Worksheet
Dim LastRow As Long
Dim LastColumn As Long
Dim StartCell As Range

Set sht = Worksheets("Sheet1")
Set StartCell = Range("D9")

'Refresh UsedRange
  Worksheets("Sheet1").UsedRange

'Find Last Row and Column
  LastRow = StartCell.SpecialCells(xlCellTypeLastCell).Row
  LastColumn = StartCell.SpecialCells(xlCellTypeLastCell).Column

'Select Range
  sht.Range(StartCell, sht.Cells(LastRow, LastColumn)).Select

  Set tbl = ActiveSheet.ListObjects.Add(xlSrcRange, Selection, , xlYes)
  tbl.Name = "Clinical_Table"
  tbl.TableStyle = "TableStyleLight12"

End Sub
