Sub ArrayTest()

  Dim sht As Worksheet
  Dim LastRow As Long
  Dim LastColumn As Long
  Dim StartCell As range
  Dim Sheet As Worksheet
  Dim Sheet_Name As String
  Dim WkNames As Variant
  Dim TblNames As Variant

  Application.Calculation = xlCalculationManual
  Application.EnableEvents = False

  WkNames = Array("Results", "Validation Sheet", "Validated Mappings")
  TblNames = Array("Results_Tbl", "Val_Tbl", "Mappings_Tbl")

  For i = 0 To UBound(WkNames)

    Sheets(WkNames(i)).Select

    Set sht = Worksheets(WkNames(i)) 'Sets value
    Set StartCell = range("A1") 'Start cell used to determine where to begin creating the table range

    'Find Last Row and Column

    LastRow = StartCell.SpecialCells(xlCellTypeLastCell).Row
    LastColumn = StartCell.SpecialCells(xlCellTypeLastCell).Column
    Sheet_Name = WkNames(i) 'Assigns sheet name to a variable as a string

    'Select Range
    sht.range(StartCell, sht.Cells(LastRow, LastColumn)).Select

    'Creates the table
    Set tbl = ActiveSheet.ListObjects.Add(xlSrcRange, Selection, , xlYes)
    tbl.Name = TblNames(i) 'Names the table
    tbl.TableStyle = "TableStyleLight12" 'Sets table color theme
    Columns.AutoFit 'Autofits columns on sheet

  Next

End Sub
