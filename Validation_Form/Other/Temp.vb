Sub Nomenclature_Auto()

Dim wb As Workbook
Dim Table_Obj As ListObject
Dim Table_ObjIsVisible As Boolean
Dim Visible_Rows_Count As Integer
Dim Results_Range As range
Dim Val_Vis_Row As range
Dim StartCell As range

'DEBUG
Code = "2556509"

'This disables settings to improve macro performance.
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Application.EnableEvents = False



  Sheets("Results").Select
  ActiveSheet.ListObjects("Results").range.AutoFilter Field:=6, Criteria1:= _
      Code, Operator:=xlAnd

  Set Results_Range = range("Results_Tbl")

  'Count number of visible rows on the Results sheet
    Visible_Rows_Count = Results_Range.SpecialCells(xlCellTypeVisible).Rows.Count

  Sheets("Validation Sheet").Select
  ActiveSheet.ListObjects("Val_Tbl").range.AutoFilter Field:=9, Criteria1:= _
      Code, Operator:=xlAnd

  'Range("A1").Select
  'Range(Selection, Selection.End(xlToRight)).Select
  'Range(Selection, Selection.End(xlDown)).Select

  Set StartCell = range("A1")

  Validation_Visible_Row = StartCell.SpecialCells(xlCellTypeLastCell).Row

  Rows(Validation_Visible_Row).Select

  Selection.copy

  Sheets("New Lines").Select

  'Used to determine sheet location when replacing nomenclature values after new lines have been created.
  Code_Blank_Line = range("N" & Rows.Count).End(xlUp).Row + 1

  'Creates a new line for each "hit" for a specific code.
  For i = 1 To Visible_Rows_Count
    'Used to determine next blank line for copying the new validation line.
    Next_Blank_Row = range("N" & Rows.Count).End(xlUp).Row + 1

    range("A" & Next_Blank_Row).Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
    :=False, Transpose:=False

  Next i

  Sheets("Results").Select

  'Confirms active cell is within the table
  range("A2").Select
  'Selects the first visible cell in column '12'
  ActiveSheet.AutoFilter.range.Offset(1).SpecialCells(xlCellTypeVisible).Cells(1, 12).Select

  range(Selection, Selection.End(xlDown)).Select
  Selection.copy

  Sheets("New Lines").Select

  range("N" & Code_Blank_Line).Select
  Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
  :=False, Transpose:=False

  'Switches back to Results Sheet to Copy next column
  Sheets("Results").Select
  'Confirms active cell is within the table
  range("A2").Select
  'Selects the first visible cell in column '11'
  ActiveSheet.AutoFilter.range.Offset(1).SpecialCells(xlCellTypeVisible).Cells(1, 11).Select

  range(Selection, Selection.End(xlDown)).Select
  Selection.copy

  Sheets("New Lines").Select

  range("O" & Code_Blank_Line).Select
  Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
  :=False, Transpose:=False

End Sub
