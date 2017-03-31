Sub Nomenclature_Notes()

    Dim sht As Worksheet
    Dim LastRow As Long
    Dim LastColumn As Long
    Dim StartCell As Range
    Dim Sheet As Worksheet
    Dim rList As Range
    Dim Confirm_Run As Integer

    'Disables settings to improve performance
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False

    'Prompts user to confirm they have reviewed the data in the validation form BEFORE running this.
    Confirm_Run = MsgBox("This program will populate the notes and team fields for the nomenclature data. Click ""Ok"" to run or ""Cancel"" to cancel the program.", vbOkCancel + vbQuestion, "Empty Sheet")

    'If user hits cancel then close program.
    If Confirm_Run = vbCancel Then
      MsgBox ("Program is canceling per user action.")
      Exit Sub
    End If

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

      Rows("1:1").Select
      With Selection.Font
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = 0
      End With


      ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
      '           Populates - Event Code and Nomenclature are not mapped
      ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

      'Applies correct filter for this note value
      ActiveSheet.ListObjects("New_Lines").Range.AutoFilter Field:=5, Criteria1:= _
          "=PowerForm", Operator:=xlOr, Criteria2:="=IView"
      ActiveSheet.ListObjects("New_Lines").Range.AutoFilter Field:=13, Criteria1:= _
          Array("Alpha List", "Alpha Combo", "Discrete Grid", "UltraGrid", "PowerGrid", "Multi"), Operator:=xlFilterValues
      ActiveSheet.ListObjects("New_Lines").Range.AutoFilter Field:=19, Criteria1:= _
          "0"
      ActiveSheet.ListObjects("New_Lines").Range.AutoFilter Field:=20, Criteria1:= _
          "0"

      '''Loops Through filtered rows''''
      Range("Q3:Q" & Cells.SpecialCells(xlCellTypeLastCell).Row).Select
      Selection.Name = "Visible_Range"

      'For each visible cell within Range'
      For Each cell In Range("Visible_Range").SpecialCells(xlCellTypeVisible)
          cell.Value = "This nomenclature and event code are not mapped and should be if this will be used to complete the measure."
          cell.Offset(0,1).value = "PCST"
      Next



      '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
      '         Populates - nomenclature is mapped, but the Event Code is Not
      '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

      'Applies correct filter for this note value
      ActiveSheet.ListObjects("New_Lines").Range.AutoFilter Field:=5, Criteria1:= _
          "=PowerForm", Operator:=xlOr, Criteria2:="=IView"
      ActiveSheet.ListObjects("New_Lines").Range.AutoFilter Field:=13, Criteria1:= _
          Array("Alpha List", "Alpha Combo", "Discrete Grid", "UltraGrid", "PowerGrid", "Multi"), Operator:=xlFilterValues
      ActiveSheet.ListObjects("New_Lines").Range.AutoFilter Field:=19, Criteria1:= _
          "0"
      ActiveSheet.ListObjects("New_Lines").Range.AutoFilter Field:=20, Criteria1:= _
          "Validated"

      '''Loops Through filtered rows'''
      Range("Q3:Q" & Cells.SpecialCells(xlCellTypeLastCell).Row).Select
      Selection.Name = "Visible_Range"

      'For each visible cell within Range'
      For Each cell In Range("Visible_Range").SpecialCells(xlCellTypeVisible)
          cell.Value = "This nomenclature is mapped but the event code will need to be mapped if this will be used to complete the measure."
          cell.Offset(0,1).value = "PCST"
      Next


      '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
      '   Populates - Event Code is mapped, but the nomenclature is not
      '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

      'Applies correct filter for this note value
      ActiveSheet.ListObjects("New_Lines").Range.AutoFilter Field:=5, Criteria1:= _
          "=PowerForm", Operator:=xlOr, Criteria2:="=IView"
      ActiveSheet.ListObjects("New_Lines").Range.AutoFilter Field:=13, Criteria1:= _
          Array("Alpha List", "Alpha Combo", "Discrete Grid", "UltraGrid", "PowerGrid", "Multi"), Operator:=xlFilterValues
      ActiveSheet.ListObjects("New_Lines").Range.AutoFilter Field:=19, Criteria1:= _
          "Validated"
      ActiveSheet.ListObjects("New_Lines").Range.AutoFilter Field:=20, Criteria1:= _
          "0"

      '''Loops Through filtered rows'''
      Range("Q3:Q" & Cells.SpecialCells(xlCellTypeLastCell).Row).Select
      Selection.Name = "Visible_Range"

      'For each visible cell within Range'
      For Each cell In Range("Visible_Range").SpecialCells(xlCellTypeVisible)
          cell.Value = "This event code is mapped but the nomenclature is not mapped and should be if this will be used to complete the measure."
          cell.Offset(0,1).value = "Consulting"
      Next

      'Re-enables previously disabled settings after all code has run.
      Application.ScreenUpdating = True
      Application.Calculation = xlCalculationAutomatic
      Application.EnableEvents = True

      MsgBox("Program Completed")

End Sub
