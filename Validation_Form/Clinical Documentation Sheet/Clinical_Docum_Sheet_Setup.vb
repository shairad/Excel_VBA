Sub Clin_Docum_Sheet_Setup()

Dim tbl As ListObject
Dim sht As Worksheet
Dim LastRow As Long
Dim LastColumn As Long
Dim StartCell As Range

TESTING AGAIN!!!!

Application.ScreenUpdating = False

Sheets("Clinical Documentation").Select

Set sht = Worksheets("Clinical Documentation")
Set StartCell = Range("A2")

'Refresh UsedRange
  Worksheets("Clinical Documentation").UsedRange

'Find Last Row and Column
  LastRow = StartCell.SpecialCells(xlCellTypeLastCell).Row
  LastColumn = StartCell.SpecialCells(xlCellTypeLastCell).Column

'Select Range
sht.Range(StartCell, sht.Cells(LastRow, LastColumn)).Select

    'Turn selected Range Into Table
Set tbl = ActiveSheet.ListObjects.Add(xlSrcRange, Selection, , xlYes)
  tbl.Name = "Clinical_Table"
  tbl.TableStyle = "TableStyleLight9"

Range("A2").Select

End Sub
