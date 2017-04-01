Private Sub Summary_Sheet_Initial_Setup()
'
' Sets up the Summary Sheet. Copies the Registries and Measures and creates the concat column
'

	Dim tbl As ListObject
	Dim HeaderNames As Variant
	Dim HeaderLocations As Variant
	Dim rList As Range

	'This disables settings to improve macro performance.
	Application.ScreenUpdating = False
	Application.Calculation = xlCalculationManual
	Application.EnableEvents = False

	HeaderNames = Array("Registry", "Measure", "Concat")
	SummaryColumns = Array("Reg", "Meas", "Concat")

	Sheets("Summary View").Select

	ActiveSheet.AutoFilterMode = False

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

'Clears formats
	Range("B1").Select
	Range(Selection, Selection.End(xlToRight)).Select
	Selection.ClearFormats

'Clear cell values if there are analysis
	Sheets("Summary View").Select
	Range("A1:K1").Select
	Selection.Name = "Summary_Headers"

	For Each cell In Range("Summary_Headers")
		CurrentHeader = cell

		IsInHeaderArray = Not IsError(Application.Match(CurrentHeader, HeaderNames, 0))

		If IsInHeaderArray = True Then
			CurrentAddress = Mid(cell.Address, 2, 1)
			Range(CurrentAddress & "2").Select
			Range(Selection, Selection.End(xlDown)).Select
			Selection.Clear
		End If
		Next cell



		Sheets("Summary View").Select
		Range("B1:J1").Select
		Selection.Name = "Header_Row"

'finds column letter for each of the colums we care about
		For Each cell In Range("Header_Row")

			If cell = "Registry" Then
				SummaryColumns(0) = Mid(cell.Address, 2, 1)

			ElseIf cell = "Measure" Then
				SummaryColumns(1) = Mid(cell.Address, 2, 1)

			ElseIf cell = "Concat" Then
				SummaryColumns(2) = Mid(cell.Address, 2, 1)

'Elseif cell = "Health Maintenance" Then
'  SummaryColumns(3) = Mid(cell.Address, 2, 1)
			End If

			Next cell


'Copies the Registry and measure columns to the summary view sheet
			Sheets("Combined Registry Measures").Select
			Range("A2:B2").Select
			Range(Selection, Selection.End(xlDown)).Select
			Selection.Copy
			Sheets("Summary View").Select

'Uses the location of the Registry column to paste the data
			Range(SummaryColumns(0) & "2").Select
			Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
			:=False, Transpose:=False


			Range("B1").Select
			Range(Selection, Selection.End(xlToRight)).Select
			Range(Selection, Selection.End(xlDown)).Select


			Set tbl = ActiveSheet.ListObjects.Add(xlSrcRange, Selection, , xlYes)
			tbl.Name = "Summary_Table"
			tbl.TableStyle = "TableStyleLight13"

'Changes header font back to white
			Rows("1:1").Select
			With Selection.Font
				.ThemeColor = xlThemeColorDark1
				.TintAndShade = 0
			End With

'Uses the location of the concat column
			Range(SummaryColumns(2) & "2").Select
			ActiveCell.Formula = "=CONCATENATE(B2,""|"",C2)"


'Re-enables previously disabled settings after all code has run.
			Application.ScreenUpdating = True
			Application.Calculation = xlCalculationAutomatic
			Application.EnableEvents = True



		End Sub
