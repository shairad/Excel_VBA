Sub Name_Current_Selection_As_Range()
'
' Selects all cells on the sheet and converts them into a named table
'

Dim lastrow As Long
Dim sht As Worksheet
Dim LastColumn As Long
Dim StartCell As Range


    Set sht = Worksheets("Clinical Documentation")
  	Set StartCell = Range("A2")

  'Refresh UsedRange
  	Worksheets("Clinical Documentation").UsedRange

  'Find Last Row and Column
  	LastRow = StartCell.SpecialCells(xlCellTypeLastCell).Row
  	LastColumn = StartCell.SpecialCells(xlCellTypeLastCell).Column

  'Select Range
  	sht.Range(StartCell, sht.Cells(LastRow, LastColumn)).Select

    ActiveWorkbook.Names.Add Name:="Test", RefersTo:= Selection 'Names current selection



  End Sub
