Sub Summary_Sheet_Setup()

	answer = MsgBox("This will launch the Summary sheet scripts. Leave computer alone until completed." & vbNextLine & "Are you ready?", vbYesNo + vbQuestion, "Empty Sheet")

	If answer = vbYes Then

	Application.ScreenUpdating = False

		Call Unmapped_Summary_Pivot
		Call Validated_Summary_Pivot
		Call Clinical_Summary_Sheet_Pivot
		Call Summary_Combined_Lookup_Sheet
		Call Summary_Sheet_Initial_Setup
		Call Summary_Add_HyperLinks


	Else
'do nothing
	End If

	Application.ScreenUpdating = True

End Sub
