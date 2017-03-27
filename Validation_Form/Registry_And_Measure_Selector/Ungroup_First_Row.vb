Private Sub ungroup_first_row()

' Group Macro creates a one line group with the header file. This macro deletes that grouping to 'keep things clean.

	Sheets("Pivot").Select
	Range("A1").Select
	Selection.Rows.Ungroup
End Sub
