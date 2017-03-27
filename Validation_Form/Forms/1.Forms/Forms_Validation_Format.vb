Sub Validation_Format()
'
'Loops through the table and deletes all the extra header rows


	Dim RowToTest As Long
	Dim tbl As ListObject
	Dim Rng As Range
	Dim rList As Range

	Application.ScreenUpdating = False

'Converts numbers stored as text to number format

	Sheets("Validated Codes").Select 'Selects Sheet
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

	Range("D1:D" & ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row).Select 'Selects all cells not empty in column
	Selection.Name = "Validated_Code_ID" 'Names Range

	Set Rng = Range("Validated_Code_ID") 'Assigns range to variable

	For Each cell In Rng 'Loops through cells in range

		If IsNumeric(cell) Then 'If cell contains numbers then X
			cell.Select 'Select the cell
			With Selection 'With the selected cell convert cell format to number without any decimal places
				Selection.NumberFormat = "0"
				.Value = .Value
			End With

		End If
		Next cell


'Removes extra formatting of cells and standardizes all cells in the same format then formats 'range as table.

		Sheets("Validation Format").Select
		Cells.Select
		Selection.Style = "Normal"
		ActiveSheet.AutoFilterMode = False 'Turns off autofilter

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

		Range("I1").Select
		Range(Selection, Selection.End(xlDown)).Select
		Range(Selection, Selection.End(xlToLeft)).Select
		Set tbl = ActiveSheet.ListObjects.Add(xlSrcRange, Selection, , xlYes)
		tbl.Name = "Forms_Val"
		tbl.TableStyle = "TableStyleLight12"



		For RowToTest = Cells(Rows.Count, 2).End(xlUp).Row To 2 Step -1

			With Cells(RowToTest, 2)
				If .Value = "FORM_DEFINITION" Then
					Rows(RowToTest).EntireRow.delete
				End If
			End With
			Next RowToTest


			Range("A2").Select
			ActiveCell = _
			"=IFERROR(INDEX('Validated Codes'!I:I,MATCH(E2,'Validated Codes'!D:D,0)),0)"
			Selection.AutoFill Destination:=Range("Forms_Val[Mapped?]")

			Columns("A:A").Select
			With Selection
				.HorizontalAlignment = xlCenter
				.VerticalAlignment = xlBottom
				.WrapText = False
				.Orientation = 0
				.AddIndent = False
				.IndentLevel = 0
				.ShrinkToFit = False
				.ReadingOrder = xlContext
				.MergeCells = False
			End With
			Range("A2").Select

			Application.ScreenUpdating = True

		End Sub
