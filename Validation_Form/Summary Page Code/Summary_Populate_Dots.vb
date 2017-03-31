Sub Summary_Pop_Dots()
'
' Copies the temporary values from the lookup columns and pastes the VALUES into the appropriate columns.

	answer = MsgBox("This will populate the street lights on the Summary sheet. Leave computer alone until completed." & vbNextLine & "Are you ready?", vbYesNo + vbQuestion, "Empty Sheet")

	If answer = vbYes Then

		Application.ScreenUpdating = False

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

		Sheets("Clinical_Summary_Pivot").Select 'Confirms Clinical Summary Pivot is up to date
		ActiveSheet.PivotTables("Clinical_Summary_Pivot").PivotCache.Refresh

		Sheets("Validated_Summary_Pivot").Select 'Confirms Validation Summary Pivot is up to date
		ActiveSheet.PivotTables("Validated_Pivot").PivotCache.Refresh

		Sheets("Unmapped_Summary_Pivot").Select 'Confirms Unmapped Summary Pivot is up to date
		ActiveSheet.PivotTables("Unmapped_Pivot").PivotCache.Refresh


		''Selects the Validated Lookup Values
		Sheets("Combined Registry Measures").Select
		Range("E2").Select
		Range(Selection, Selection.End(xlDown)).Select
		Selection.Copy



		''''''OLD''''''
		Sheets("Combined Registry Measures").Select
		Range("E2:G2").Select
		Range(Selection, Selection.End(xlDown)).Select
		Selection.Copy
		Sheets("Summary View").Select
		Range("E2").Select
		Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
		:=False, Transpose:=False

		Range("B2").Select
		Range("B2:K2").Select
		Range(Selection, Selection.End(xlDown)).Select
		Selection.ClearFormats

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

	Else

'Do Nothing

	End If

	Sheets("Summary View").Select
	Range("E2").Select
	Range(Selection, Selection.End(xlDown)).Select
	ActiveSheet.Hyperlinks.Add Anchor:=Selection, Address:="", SubAddress:= _
	"'Potential Mapping Issues'!A1"

	Range("F2").Select
	Range(Selection, Selection.End(xlDown)).Select
	ActiveSheet.Hyperlinks.Add Anchor:=Selection, Address:="", SubAddress:= _
	"'Unmapped Codes'!A1"

	Range("G2").Select
	Range(Selection, Selection.End(xlDown)).Select
	ActiveSheet.Hyperlinks.Add Anchor:=Selection, Address:="", SubAddress:= _
	"'Clinical Documentation'!A1"


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

	Application.ScreenUpdating = True

End Sub
