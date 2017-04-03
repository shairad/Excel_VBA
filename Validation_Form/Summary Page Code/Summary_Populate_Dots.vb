Private Sub Summary_Pop_Dots()
'
' Copies the temporary values from the lookup columns and pastes the VALUES into the appropriate columns.

	Dim Sheet_Headers As Variant
	Dim Find_Header As Range
	Dim rngHeaders As Range
	Dim ColHeaders As Variant
	Dim Validated_Col As Variant
	Dim Unmapped_Col As Variant
	Dim Clinical_Col As Variant
	Dim Health_Col As Variant
	Dim WkNames As Variant
	Dim PivotNames As Variant
	Dim CopyColumns As Variant
	Dim SummaryColumns As Variant
	Dim HyperLinkSheets As Variant


	'This disables settings to improve macro performance.
	Application.ScreenUpdating = False
	Application.Calculation = xlCalculationManual
	Application.EnableEvents = False

	Sheets("Summary View").Select
	Columns("E:H").Select
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

	'Refreshes pivot table data
	WkNames = Array("Potential_Summary_Pivot", "Clinical_Summary_Pivot", "Unmapped_Summary_Pivot")
	PivotNames = Array("Potential_Pivot", "Clinical_Pivot", "Unmapped_Pivot")
	CopyColumns = Array("E2", "F2", "G2")
	SummaryColumns = Array("False", "False", "False")
	HyperLinkSheets = Array("'Potential Mapping Issues'", "'Clinical Documentation'" , "'Unmapped Codes'")

	For i = 0 To UBound(WkNames)

		CurrentWk =WkNames(i)
		Sheets(CurrentWk).Select
		ActiveSheet.PivotTables(PivotNames(i)).PivotCache.Refresh

	Next i


	''''''finds and stores summary header columns''''''''
	Sheets("Summary View").Select
  Range("B1:J1").Select
  Selection.Name = "Header_Row"


	'finds column letter for each of the colums we care about
  For each cell in Range("Header_Row")

    If cell = "Potential Mapping Issues" Then
      SummaryColumns(0) = Mid(cell.Address, 2, 1)

    Elseif cell = "Unmapped Codes" Then
      SummaryColumns(1) = Mid(cell.Address, 2, 1)

    Elseif cell = "Clinical Documentation" Then
      SummaryColumns(2) = Mid(cell.Address, 2, 1)

    'Elseif cell = "Health Maintenance" Then
    '  SummaryColumns(3) = Mid(cell.Address, 2, 1)
    End if

  Next cell



	For i = 0 to UBound(CopyColumns)
		CurrentWk = WkNames(i)
		CurrentCopyCol = CopyColumns(i)
		CurrentSumCol = SummaryColumns(i)

		'Confirms the column exists. If the column does not exist then skip it.
		If CurrentSumCol <> "False" Then

			Sheets("Combined Registry Measures").Select
			Range(CurrentCopyCol).Select
			Range(Selection, Selection.End(xlDown)).Select
			Selection.Copy

			Sheets("Summary View").Select
			Range(CurrentSumCol & "2").Select
			Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
			:=False, Transpose:=False

		End If

	Next i


	Range("B2:H2").Select
	Range(Selection, Selection.End(xlDown)).Select
	Selection.ClearFormats

	'Autofit for all cells on screen.
	Cells.Select
	Cells.EntireColumn.AutoFit

	'Applies the street light rules to the range
	Range("E2:H2").Select
	Range(Selection, Selection.End(xlDown)).Select
	Application.CutCopyMode = False
	Selection.FormatConditions.AddIconSetCondition
	Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
	With Selection.FormatConditions(1)
		.ReverseOrder = True
		.ShowIconOnly = True
		.IconSet = ActiveWorkbook.IconSets(xl3TrafficLights1)
	End With

	With Selection.FormatConditions(1).IconCriteria(2)
		.Type = xlConditionValueNumber
		.Value = 1
		.Operator = 7
	End With

	With Selection.FormatConditions(1).IconCriteria(3)
		.Type = xlConditionValueNumber
		.Value = 4
		.Operator = 7
	End With

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


	'Adds the hyperlink address to the street lights
	For i = 0 To UBound(SummaryColumns)
		CurrentWk = WkNames(i)
		CurrentCopyCol = CopyColumns(i)
		CurrentSumCol = SummaryColumns(i)
		CurrentHyperSht = HyperLinkSheets(i)

		'Confirms the column exists. If the column does not exist then skip it.
		If CurrentSumCol <> "False" Then

		Range(CurrentSumCol & "2").Select
		Range(Selection, Selection.End(xlDown)).Select
		ActiveSheet.Hyperlinks.Add Anchor:=Selection, Address:="", SubAddress:= _
		CurrentHyperSht & "!A1"

		End If

	Next i


	'Formats the angle for the header row of Summary Sheet
	Rows("1:1").Select

	With Selection
		.VerticalAlignment = xlBottom
		.WrapText = False
		.Orientation = 45
		.AddIndent = False
		.ShrinkToFit = False
		.ReadingOrder = xlContext
		.MergeCells = False
	End With

	'Autofit for all cells on screen.
	Cells.Select
	Cells.EntireColumn.AutoFit

		'Cleans up selected cells on sheet.
	Range("A1").Select

'Re-enables previously disabled settings after all code has run.
	Application.ScreenUpdating = True
	Application.Calculation = xlCalculationAutomatic
	Application.EnableEvents = True

End Sub
