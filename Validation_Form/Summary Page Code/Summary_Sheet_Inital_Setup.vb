Private Sub Summary_Sheet_Initial_Setup()
'
' Summary_Setup_1 Macro
'
	answer = MsgBox("This is the summary sheet initial setup. Only run this after all analysis has been completed. Leave computer alone until completed." & vbNextLine & "Are you ready?", vbYesNo + vbQuestion, "Empty Sheet")

	If answer = vbYes Then

		Application.ScreenUpdating = False

		Dim tbl As ListObject

		Cells.Select
		Selection.ClearFormats

		Sheets("Combined Registry Measures").Select
		Range("A2:B2").Select
		Range(Selection, Selection.End(xlDown)).Select
		Selection.Copy
		Sheets("Summary View").Select

		Range("B2").Select
		Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
		:=False, Transpose:=False

		Sheets("Summary View").Select

		Range("L1").Select
		ActiveCell.Formula = "Concat"

		Range("L2").Select
		ActiveCell.Formula = "=CONCATENATE(B2,""|"",C2)"
		Range("L3").Select

		Range("B1").Select
		Range(Selection, Selection.End(xlToRight)).Select
		Range(Selection, Selection.End(xlDown)).Select


		Set tbl = ActiveSheet.ListObjects.Add(xlSrcRange, Selection, , xlYes)
		tbl.Name = "Summary_Table"
		tbl.TableStyle = "TableStyleLight13"

		Range("L2").Select
		Application.CutCopyMode = False
		Selection.AutoFill Destination:=Range("Summary_Table[Concat]")
		Range("Summary_Table[Concat]").Select

	Else
'Do Nothing

	End If

	Application.ScreenUpdating = True

End Sub
