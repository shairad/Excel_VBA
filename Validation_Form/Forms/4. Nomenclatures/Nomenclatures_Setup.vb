Sub Nomenclature_Setup()
'
' Nomenclature_Results Macro. Formats range as table. Inserts lookup formulas and populates autofill.
'
	Dim tbl As ListObject
	Dim sht As Worksheet
	Dim LastRow As Long
	Dim LastColumn As Long
	Dim StartCell As Range

	Application.ScreenUpdating = False


	Sheets("Results").Select

	Range("E1:E" & ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row).Select 'Selects all cells not empty in column
	Selection.Name = "DTA_Code_ID" 'Names Range

	Set Rng = Range("DTA_Code_ID") 'Assigns range to variable

	For Each cell In Rng 'Loops through cells in range

		If IsNumeric(cell) Then 'If cell contains numbers then X
			cell.Select 'Select the cell
			With Selection 'With the selected cell convert cell format to number without any decimal places
				Selection.NumberFormat = "0"
				.Value = .Value
			End With

		End If
		Next cell

		ActiveSheet.AutoFilterMode = False 'Disables filters on sheet

'If Sheet contains a table, convert table to range.
		If ActiveSheet.ListObjects.Count > 0 Then

			With ActiveSheet.ListObjects(1)
				Set rList = .Range
				.Unlist                           ' convert the table back to a range
			End With

			With rList
				.Interior.ColorIndex = xlColorIndexNone
				.Font.ColorIndex = xlColorIndexAutomatic
				.Borders.LineStyle = xlLineStyleNone
			End With

		End If

		Set sht = Worksheets("Results")
		Set StartCell = Range("A1")

'Refresh UsedRange
		Worksheets("Results").UsedRange

'Find Last Row and Column
		LastRow = StartCell.SpecialCells(xlCellTypeLastCell).Row
		LastColumn = StartCell.SpecialCells(xlCellTypeLastCell).Column

'Select Range
		sht.Range(StartCell, sht.Cells(LastRow, LastColumn)).Select

'Turn selected Range Into Table
		Set tbl = ActiveSheet.ListObjects.Add(xlSrcRange, Selection, , xlYes)
		tbl.Name = "Nomenclature_Results"
		tbl.TableStyle = "TableStyleLight12"
		Range("A2").Select

		Range("A2").Select
		ActiveCell = "=IFERROR(INDEX('Validated Mappings'!I:I,MATCH(E2,'Validated Mappings'!D:D,0)),0)"
		Selection.AutoFill Destination:=Range("Nomenclature_Results[Event Code Mapped?]")
		Range("Nomenclature_Results[Event Code Mapped?]").Select

		Range("B2").Select
		ActiveCell = "=IFERROR(INDEX('Validated Mappings'!I:I,MATCH(K2,'Validated Mappings'!D:D,0)),0)"
		Selection.AutoFill Destination:=Range("Nomenclature_Results[Nomenclature Mapped?]")
		Range("Nomenclature_Results[Nomenclature Mapped?]").Select


		Sheets("Results").Select
		Cells.Select
		Selection.Copy
		Sheets("To_Review").Select
		Range("A1").Select
		Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
		:=False, Transpose:=False

		Set sht = Worksheets("To_Review")
		Set StartCell = Range("A1")

'Refresh UsedRange
		Worksheets("To_Review").UsedRange

'Find Last Row and Column
		LastRow = StartCell.SpecialCells(xlCellTypeLastCell).Row
		LastColumn = StartCell.SpecialCells(xlCellTypeLastCell).Column

'Select Range
		sht.Range(StartCell, sht.Cells(LastRow, LastColumn)).Select

		Set tbl = ActiveSheet.ListObjects.Add(xlSrcRange, Selection, , xlYes)
		tbl.Name = "Nomenclature_To_Review"
		tbl.TableStyle = "TableStyleLight9"

		Cells.Select
		Cells.EntireColumn.AutoFit
		Cells.Select
		Cells.EntireRow.AutoFit

		Columns("A:B").Select
		With Selection
			.HorizontalAlignment = xlCenter
			.VerticalAlignment = xlCenter
			.Orientation = 0
			.AddIndent = False
			.IndentLevel = 0
			.ShrinkToFit = False
			.ReadingOrder = xlContext
			.MergeCells = False
		End With

		Range("A1").Select

		Application.ScreenUpdating = True

	End Sub
