Sub Convert()


'Disables settings to improve performance
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Application.EnableEvents = False


Selection.Name = "Nomen_Code_ID" 'Names Range

Set Rng = Range("Nomen_Code_ID") 'Assigns range to variable

For Each cell In Rng 'Loops through cells in range
  If IsNumeric(cell) Then
    cell.Value = Val(cell.Value)
    cell.NumberFormat = "0"
  Elseif IsDate(cell) Then
    cell.Value = DateValue(cell.Value)
    cell.NumberFormat = "mm-dd-yyyy"
  End If

Next cell

Application.ScreenUpdating = True
Application.EnableEvents = True

End Sub
