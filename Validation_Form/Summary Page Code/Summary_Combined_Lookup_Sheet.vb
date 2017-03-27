Private Sub Summary_Combined_Lookup_Sheet()
'
' Takes the Registries, Measures and Concepts from the Unmapped and Validated Sheets and combinds them into one sheet.Then creates a CONCATENATE column for lookup.
'

	With ThisWorkbook
		.Sheets.Add(After:=.Sheets(.Sheets.Count)).Name = "Combined Registry Measures"
	End With

'Takes all the concept measures and puts them on one sheet
	Sheets("Potential Mapping Issues").Select
	Range("Validated_Mappings_Table[[#Headers],[Registry]:[Concept]]").Select
	Range(Selection, Selection.End(xlDown)).Select
	Selection.Copy
	Sheets("Combined Registry Measures").Select
	Range("A1").Select
	Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
	:=False, Transpose:=False

	Sheets("Unmapped Codes").Select
	Range("B3:D3").Select
	Range(Selection, Selection.End(xlDown)).Select
	Application.CutCopyMode = False
	Selection.Copy
	Sheets("Combined Registry Measures").Select
	Selection.End(xlDown).Select
	ActiveCell.Offset(1, 0).Range("A1").Select
	ActiveSheet.Paste

	Range("D1").Select
	ActiveCell.Formula = "Concat"

	Range("E1").Select
	ActiveCell.Formula = "Validated_lookup"

	Range("F1").Select
	ActiveCell.Formula = "Unmapped_lookup"

	Range("G1").Select
	ActiveCell.Formula = "Clinical_lookup"

	Range("A1").Select
	Range(Selection, Selection.End(xlDown)).Select
	Range(Selection, Selection.End(xlToRight)).Select
	Application.CutCopyMode = False

	Dim Ws As Worksheet
	Set Ws = ThisWorkbook.Sheets("Combined Registry Measures")

	Sheets("Combined Registry Measures").Select
	Range("A1:G" & ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row).Select

	Ws.ListObjects.Add(xlSrcRange, Selection, , xlYes).Name = "combined_lookup_range"
	Ws.ListObjects("combined_lookup_range").TableStyle = "TableStyleLight12"

	Sheets("Combined Registry Measures").Select
	ActiveSheet.Range("combined_lookup_range[#All]").RemoveDuplicates Columns:=Array(1, 2), Header:=xlYes

	Range("D2").Select
	ActiveCell.Formula = "=CONCATENATE(A2,""|"",B2)"
	Range("D3").Select
	Columns("D:D").EntireColumn.AutoFit

	Range("E2").Select
	ActiveCell.Formula = _
	"=IFERROR(INDEX(Validated_Summary_Pivot!C:C,MATCH(D2,Validated_Summary_Pivot!D:D,0)),0)"

	Range("F2").Select
	ActiveCell.Formula = _
	"=IFERROR(INDEX(Unmapped_Summary_Pivot!C:C,MATCH(D2,Unmapped_Summary_Pivot!D:D,0)),0)"

	Range("G2").Select
	ActiveCell.Formula = _
	"=IFERROR(INDEX(Clinical_Summary_Pivot!C:C,MATCH(D2,Clinical_Summary_Pivot!D:D,0)),0)"

End Sub
