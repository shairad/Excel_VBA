Sub EV_Code_Setup()

    Dim tbl As ListObject
    Dim sht As Worksheet
    Dim LastRow As Long
    Dim LastColumn As Long
    Dim StartCell As Range

    Application.ScreenUpdating = False

		Sheets("Event Codes Results").Select

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

    Set sht = Worksheets("Event Codes Results")
    Set StartCell = Range("A1")

    'Refresh UsedRange
      Worksheets("Event Codes Results").UsedRange

    'Find Last Row and Column
      LastRow = StartCell.SpecialCells(xlCellTypeLastCell).Row
      LastColumn = StartCell.SpecialCells(xlCellTypeLastCell).Column

    'Select Range
    sht.Range(StartCell, sht.Cells(LastRow, LastColumn)).Select

        'Turn selected Range Into Table
    Set tbl = ActiveSheet.ListObjects.Add(xlSrcRange, Selection, , xlYes)
    tbl.Name = "EV_Results_Table"
    tbl.TableStyle = "TableStyleLight9"

    'changes font color of header row to white
    Rows("1:1").Select
    With Selection.Font
      .ThemeColor = xlThemeColorDark1
      .TintAndShade = 0
    End With
    
    Range("A2").Select

    Sheets("Event Codes Results").Select
    Range("A2").Select
    ActiveCell = "=IFERROR(INDEX('Validated Codes'!I:I,MATCH(D3,'Validated Codes'!D:D,0)),0)"
    Selection.AutoFill Destination:=Range("EV_Results_Table[Mapped?]")
    Range("A2").Select

    'Creates and formats Pivot TableStyle

        Sheets("Event Codes Results").Select

        ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
            "EV_Results_Table", Version:=6).CreatePivotTable TableDestination:= _
            "'Pivot Table'!R12C2", TableName:="EV_Pivot", DefaultVersion:=6
        Sheets("Pivot Table").Select
        Cells(12, 2).Select
        With ActiveSheet.PivotTables("EV_Pivot").PivotFields("Mapped?")
            .Orientation = xlRowField
            .Position = 1
        End With
        With ActiveSheet.PivotTables("EV_Pivot").PivotFields("CODE_STATUS")
            .Orientation = xlColumnField
            .Position = 1
        End With
        ActiveSheet.PivotTables("EV_Pivot").AddDataField ActiveSheet.PivotTables( _
            "EV_Pivot").PivotFields("EVENT_CD"), "Count of EVENT_CD", xlCount

      'Selects the Validated count and changes color to red
      Range("C15").Select
      With Selection.Font
          .Color = -16776961
          .TintAndShade = 0
      End With


    Sheets("Event Codes Results").Select
      'Filters the results table column A for just "Validated"
    ActiveSheet.ListObjects("EV_Results_Table").Range.AutoFilter Field:=1, _
        Criteria1:="Validated"
      'Filters the Code_Status column for just "Active"
    ActiveSheet.ListObjects("EV_Results_Table").Range.AutoFilter Field:=3, _
        Criteria1:="Active"

    Application.ScreenUpdating = True

End Sub
