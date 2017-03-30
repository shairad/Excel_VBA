Sub SpecialLoop()

    Dim sht As Worksheet
    Dim LastRow As Long
    Dim LastColumn As Long
    Dim StartCell As Range
    Dim Sheet As Worksheet
    Dim rList As Range

    ActiveSheet.AutoFilterMode = False 'Removes filters from sheet

    If ActiveSheet.ListObjects.Count > 0 Then

    With ActiveSheet.ListObjects(1)
        Set rList = .Range
        .Unlist                           ' convert the table back to a range
    End With

    With rList
        .Interior.ColorIndex = xlColorIndexNone
        .Font.ColorIndex = xlColorIndexAutomatic
        .Borders.LineStyle = xlLineStyleNone
    End With

    Rows("1:1").Select
    With Selection.Font
      .ThemeColor = xlThemeColorDark1
      .TintAndShade = 0
    End With

    End If


      Set sht = ActiveSheet 'Sets value
      Set StartCell = Range("A1") 'Start cell used to determine where to begin creating the table range

    'Find Last Row and Column
      LastRow = StartCell.SpecialCells(xlCellTypeLastCell).Row
      LastColumn = StartCell.SpecialCells(xlCellTypeLastCell).Column

    'Select Range
      sht.Range(StartCell, sht.Cells(LastRow, LastColumn)).Select

    'Creates the table
      Set tbl = ActiveSheet.ListObjects.Add(xlSrcRange, Selection, , xlYes)
      tbl.Name = "New_Lines" 'Names the table
      tbl.TableStyle = "TableStyleLight12" 'Sets table color theme


      '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
      ' Populates "This nomenclature is mapped but the event code will need to be mapped if this will be used to complete the measure."
      '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

      'Applies correct filter for this note value
      ActiveSheet.ListObjects("New_Lines").Range.AutoFilter Field:=5, Criteria1:= _
          "=PowerForm", Operator:=xlOr, Criteria2:="=IView"
      ActiveWindow.SmallScroll Down:=-6
      ActiveSheet.ListObjects("New_Lines").Range.AutoFilter Field:=13, Criteria1:= _
          Array("Alpha List", "Alpha Combo", "Discrete Grid", "UltraGrid", "PowerGrid", "Multi"), Operator:=xlFilterValues
      ActiveSheet.ListObjects("New_Lines").Range.AutoFilter Field:=19, Criteria1:= _
          "0"
      ActiveSheet.ListObjects("New_Lines").Range.AutoFilter Field:=20, Criteria1:= _
          "0"


      '''Loops Through filtered rows'''''
      Range("R3:R" & Cells.SpecialCells(xlCellTypeLastCell).Row).Select
      Selection.Name = "Visible_Range"

      'For each visible cell within Range'
      For Each cell In Range("Visible_Range").SpecialCells(xlCellTypeVisible)
          cell.Value = "This nomenclature is mapped but the event code will need to be mapped if this will be used to complete the measure."
      Next



      '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
      ' Populates "This nomenclature is mapped but the event code will need to be mapped if this will be used to complete the measure."
      '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

      'Applies correct filter for this note value
      ActiveSheet.ListObjects("New_Lines").Range.AutoFilter Field:=5, Criteria1:= _
          "=PowerForm", Operator:=xlOr, Criteria2:="=IView"
      ActiveWindow.SmallScroll Down:=-6
      ActiveSheet.ListObjects("New_Lines").Range.AutoFilter Field:=13, Criteria1:= _
          Array("Alpha List", "Alpha Combo", "Discrete Grid", "UltraGrid", "PowerGrid", "Multi"), Operator:=xlFilterValues
      ActiveSheet.ListObjects("New_Lines").Range.AutoFilter Field:=19, Criteria1:= _
          "0"
      ActiveSheet.ListObjects("New_Lines").Range.AutoFilter Field:=20, Criteria1:= _
          "0"


      '''Loops Through filtered rows'''''
      Range("R3:R" & Cells.SpecialCells(xlCellTypeLastCell).Row).Select
      Selection.Name = "Visible_Range"

      'For each visible cell within Range'
      For Each cell In Range("Visible_Range").SpecialCells(xlCellTypeVisible)
          cell.Value = "This nomenclature is mapped but the event code will need to be mapped if this will be used to complete the measure."
      Next

End Sub
