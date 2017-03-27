Sub Potential_Issues_Sheet_Setup()

Dim tbl As ListObject
Dim sht As Worksheet
Dim LastRow As Long
Dim LastColumn As Long
Dim StartCell As Range

Application.ScreenUpdating = False

Sheets("Potential Mapping Issues").Select

Set sht = Worksheets("Potential Mapping Issues")
Set StartCell = Range("A2")

'Refresh UsedRange
  Worksheets("Potential Mapping Issues").UsedRange

'Find Last Row and Column
  LastRow = StartCell.SpecialCells(xlCellTypeLastCell).Row
  LastColumn = StartCell.SpecialCells(xlCellTypeLastCell).Column

'Select Range
sht.Range(StartCell, sht.Cells(LastRow, LastColumn)).Select

    'Turn selected Range Into Table
Set tbl = ActiveSheet.ListObjects.Add(xlSrcRange, Selection, , xlYes)
  tbl.Name = "Potential_Issus_Table"
  tbl.TableStyle = "TableStyleLight9"

Range("A2").Select

End Sub
