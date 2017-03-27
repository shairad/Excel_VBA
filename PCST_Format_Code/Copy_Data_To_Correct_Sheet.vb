Sub Create_Sheets_Pop_Unmapped()

	Dim Code_Sheet As String

	For Each code In Range("Code_ID_List")

		With ThisWorkbook
			.Sheets.Add(After:=.Sheets(.Sheets.Count)).Name = code
		End With

		Sheets("Unmapped Codes").Select
		ActiveSheet.ListObjects("Unmapped_Table").range.AutoFilter Field:=12, _
		Criteria1:=code, Operator:=xlAnd
		range("Unmapped_Table[[#Headers],[Status]]").Select
		range(Selection, Selection.End(xlToRight)).Select
		range(Selection, Selection.End(xlDown)).Select
		Selection.Copy
		Code_Sheet = Code
		Sheets(Code_Sheet).Select
		Range("A1").Select
		ActiveSheet.Paste

		Next code


	End Sub
