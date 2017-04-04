Private Sub Remove_Table_Format()

	Dim rList As Range
	Dim WkNames As Variant


	WkNames = Array("Potential Mapping Issues", "Unmapped Codes", "Clinidal Documentation")

	For i = 0 to UBound(WkNames)

		On Error GoTo NoSheet
		Sheets(WkNames(i)).Select

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

		'Error handling incase sheet does not exist
		NoSheet:
			'MsgBox("No Code for " & EventCode)
			Resume ClearError

		ClearError:
		'Clears variables for next loop

	Next i

End Sub
