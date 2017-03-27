Sub Check_Number()

Dim Temp_Cell As Variant
Dim Rng As Range



Sheets("Validated Codes").Select

Range("D1:D" & ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row).Select
Selection.Name = "Validated_Code_ID"

Set Rng = Range("Validated_Code_ID")

For Each cell In Rng

      If IsNumeric(cell) Then
          Temp_Cell = cell.Value
          cell.ClearContents
          cell.Value = Temp_Cell.value

      End If
Next cell


End Sub
