Sub TimeTest()
' 100 million random numbers, tests, and math operations

Dim x As Long
Dim StartTime As Single
Dim i As Long

x = 0
StartTime = Timer

For i = 1 To 100000000
	If Rnd <= 0.5 Then
		x = x + 1
	Else x = x - 1
	End If
Next i

MsgBox(Timer - StartTime & " seconds")

End Sub
