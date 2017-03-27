Private Sub Remove_Table_Format()

	Dim rList As Range

	Sheets("Potential Mapping Issues").Select

	If ActiveSheet.ListObjects.Count > 0 Then

		With ActiveSheet.ListObjects(1)
			Set rList = .Range
			.Unlist                           ' convert the table back to a range
		End With

	End If

	If Not ActiveSheet.AutoFilterMode Then  'Adds the filter buttons to the sheet
		ActiveSheet.Range("2:2").AutoFilter
	End If

	Sheets("Unmapped Codes").Select

	If ActiveSheet.ListObjects.Count > 0 Then

		With ActiveSheet.ListObjects(1)
			Set rList = .Range
			.Unlist                           ' convert the table back to a range
		End With

	End If

	If Not ActiveSheet.AutoFilterMode Then  'Adds the filter buttons to the sheet
		ActiveSheet.Range("2:2").AutoFilter
	End If

	Sheets("Clinical Documentation").Select

	If ActiveSheet.ListObjects.Count > 0 Then

		With ActiveSheet.ListObjects(1)
			Set rList = .Range
			.Unlist                           ' convert the table back to a range
		End With

	End If

	If Not ActiveSheet.AutoFilterMode Then  'Adds the filter buttons to the sheet
		ActiveSheet.Range("2:2").AutoFilter
	End If

	Range("A2").Select


End Sub
