Sub SHX_Setup()
'
' SHX REsults Macro. Formats range as table. Inserts lookup formulas and populates autofill.
'
	Dim tbl As ListObject
	Dim sht As Worksheet
	Dim LastRow As Long
	Dim LastColumn As Long
	Dim StartCell As Range
	Dim rList As Range

	Application.ScreenUpdating = False


	Sheets("Social History Results").Select

	ActiveSheet.AutoFilterMode = False 'Disables autoFilter

'If table exists on sheet then convert to range
	If ActiveSheet.ListObjects.Count > 0 Then

		With ActiveSheet.ListObjects(1)
			Set rList = .Range
			.Unlist
		End With

		With rList
			.Interior.ColorIndex = xlColorIndexNone
			.Font.ColorIndex = xlColorIndexAutomatic
			.Borders.LineStyle = xlLineStyleNone
		End With

	End If

	Set sht = Worksheets("Social History Results")
	Set StartCell = Range("A1")

'Refresh UsedRange
	Worksheets("Social History Results").UsedRange

'Find Last Row and Column
	LastRow = StartCell.SpecialCells(xlCellTypeLastCell).Row
	LastColumn = StartCell.SpecialCells(xlCellTypeLastCell).Column

'Select Range
	sht.Range(StartCell, sht.Cells(LastRow, LastColumn)).Select

'Turn selected Range Into Table
	Set tbl = ActiveSheet.ListObjects.Add(xlSrcRange, Selection, , xlYes)
	tbl.Name = "SHX_Results"
	tbl.TableStyle = "TableStyleLight12"

	Sheets("Social History Results").Select

	Range("I1:I" & ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row).Select 'Selects all cells not empty in column
	Selection.Name = "Nomen_Code_ID" 'Names Range

	Set Rng = Range("Nomen_Code_ID") 'Assigns range to variable

	For Each cell In Rng 'Loops through cells in range

		If IsNumeric(cell) Then 'If cell contains numbers then X
			cell.Select 'Select the cell
			With Selection 'With the selected cell convert cell format to number without any decimal places
				Selection.NumberFormat = "0"
				.Value = .Value
			End With

		End If
		Next cell

		Range("A2").Select
		ActiveCell = "=IFERROR(INDEX('Validated Mappings'!I:I,MATCH(I2,'Validated Mappings'!D:D,0)),0)"
		Selection.AutoFill Destination:=Range("SHX_Results[Nomenclature Mapped?]")
		Range("SHX_Results[Nomenclature Mapped?]").Select

		Range("B2").Select
		ActiveCell = "=IFERROR(INDEX('Validated Mappings'!I:I,MATCH(F2,'Validated Mappings'!D:D,0)),0)"
		Selection.AutoFill Destination:=Range("SHX_Results[CS 72 Mapped?]")
		Range("SHX_Results[CS 72 Mapped?]").Select

		Range("C2").Select
		ActiveCell = "=IFERROR(INDEX('Validated Mappings'!I:I,MATCH(K2,'Validated Mappings'!D:D,0)),0)"
		Selection.AutoFill Destination:=Range("SHX_Results[CS 14003 Mapped?]")
		Range("SHX_Results[CS 14003 Mapped?]").Select

		Range("D2").Select
		ActiveCell = "=IFERROR(INDEX('Validated Mappings'!I:I,MATCH(M2,'Validated Mappings'!D:D,0)),0)"
		Selection.AutoFill Destination:=Range("SHX_Results[CS 4002165 Mapped?]")
		Range("SHX_Results[CS 4002165 Mapped?]").Select


		Sheets("Social History Results").Select
		Cells.Select
		Selection.Copy
		Sheets("To_Review").Select
		Range("A1").Select
		Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
		:=False, Transpose:=False

		ActiveSheet.AutoFilterMode = False

'If table exists on this sheet, then convert to range
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


		Set sht = Worksheets("To_Review")
		Set StartCell = Range("A1")

'Refresh UsedRange
		Worksheets("Social History Results").UsedRange

'Find Last Row and Column
		LastRow = StartCell.SpecialCells(xlCellTypeLastCell).Row
		LastColumn = StartCell.SpecialCells(xlCellTypeLastCell).Column

'Select Range
		sht.Range(StartCell, sht.Cells(LastRow, LastColumn)).Select

		Set tbl = ActiveSheet.ListObjects.Add(xlSrcRange, Selection, , xlYes)
		tbl.Name = "SHX_To_Review"
		tbl.TableStyle = "TableStyleLight9"

		Cells.Select
		Cells.EntireColumn.AutoFit
		Cells.Select
		Cells.EntireRow.AutoFit

		Columns("A:D").Select
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

		Application.ScreenUpdating = True 'Enables screen updating

	End Sub
