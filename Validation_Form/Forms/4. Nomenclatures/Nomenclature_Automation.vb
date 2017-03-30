Sub Nomenclature_Auto()

  Dim wb As Workbook
  Dim Table_Obj As ListObject
  Dim Table_ObjIsVisible As Boolean
  Dim Visible_Rows_Count As Integer
  Dim Confirm_Run As Integer
  Dim Results_Range As range
  Dim Val_Vis_Row As range
  Dim StartCell As range
  Dim WkNames As Variant
  Dim TblNames As Variant


  'DEBUG

  'This disables settings to improve macro performance.
  Application.ScreenUpdating = False
  Application.Calculation = xlCalculationManual
  Application.EnableEvents = False

  Confirm_Run = MsgBox("The program is about to run. This will take roughly 2 minutes. Please verify before running that you have entered all the needed data per the automation Instructions. If you have not please click cancel. Else click OK to run. ", vbOkCancel + vbQuestion, "Empty Sheet")

  'If user hits cancel then close program.
	If Confirm_Run = vbCancel Then
		MsgBox ("Program is canceling per user action.")
		Exit Sub
	End If

  WkNames = Array("Validated Mappings", "Results", "Validation Sheet")
  TblNames = Array("Mappings_Tbl", "Results_Tbl", "Val_Tbl")


  For i = 0 To UBound(WkNames)

    Sheets(WkNames(i)).Select

    If ActiveSheet.AutoFilterMode = True Then
      ActiveSheet.AutoFilterMode = False
    End If

    'Checks the current sheet. If it is in table format, convert it to range.
    If ActiveSheet.ListObjects.Count > 0 Then
      With ActiveSheet.ListObjects(1)
        Set rList = .Range
        .Unlist
      End With
      'Reverts the color of the range back to standard.
      With rList
        .Interior.ColorIndex = xlColorIndexNone
        .Font.ColorIndex = xlColorIndexAutomatic
        .Borders.LineStyle = xlLineStyleNone
      End With
    End If

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


    '''''Handles additional tasks for the Results Sheet'''''''

    If WkNames(i) = "Results" Then

    'Remove Duplicates from the Results Sheet
    ActiveSheet.range("Results_Tbl[#All]").RemoveDuplicates Columns:=Array(6, 11, 12), _
        Header:=xlYes

      'Adds new Mapping note column
      Range("M1").Select
      Selection = "Mapping Note"

      Range("A2").Select
      Selection.Formula = "=IFERROR(INDEX('Validated Mappings'!I:I,MATCH(F2,'Validated Mappings'!D:D,0)),0)"

      Range("B2").Select
      Selection.Formula = "=IFERROR(INDEX('Validated Mappings'!I:I,MATCH(L2,'Validated Mappings'!D:D,0)),0)"

      Range("C2").Select
      Selection.Formula = "=IF(AND(A2 =""Validated"", B2 = ""Validated""),""Both Validated"", 0)"

      'Hides rows which are validated in both columns
      ActiveSheet.ListObjects("Results_Tbl").range.AutoFilter Field:=3, Criteria1:= _
        "0"

    End If


    '''''''If Sheet Is The Validation Sheet'''''''''

    If WkNames(i) = "Validation Sheet" Then

      'Removes Duplicates by Event Code from the Validation Sheet
      ActiveSheet.range("Val_Tbl[#All]").RemoveDuplicates Columns:=9, Header:= _
        xlYes

      'Creates and names the Event_Codes Range which is used in loop
      Range("I2").Select
      Range(Selection, Selection.End(xlDown)).Select
      Selection.Name = "Event_Codes"
    End If

  Next i


  For Each EventCode in Range("Event_Codes")

    Sheets("Results").Select
    ActiveSheet.ListObjects("Results_Tbl").range.AutoFilter Field:=6, Criteria1:= _
        EventCode, Operator:=xlAnd

    Set Results_Range = range("Results_Tbl")

    'Error handling. If no codes are found, then skip the code.
    On Error GoTo NoBlanks

    'Count number of visible rows on the Results sheet
    Visible_Rows_Count = Results_Range.SpecialCells(xlCellTypeVisible).Rows.Count

    Sheets("Validation Sheet").Select
    ActiveSheet.ListObjects("Val_Tbl").range.AutoFilter Field:=9, Criteria1:= _
        EventCode, Operator:=xlAnd

    Set StartCell = range("A1")

    Validation_Visible_Row = StartCell.SpecialCells(xlCellTypeLastCell).Row

    Rows(Validation_Visible_Row).Select

    Selection.copy

    Sheets("New Lines").Select

    'Used to determine sheet location when replacing nomenclature values after new lines have been created.
    Code_Blank_Line = range("A" & Rows.Count).End(xlUp).Row + 1


    '''''''''''Creates a new line for each "hit" for a specific code.'''''''''''

    For i = 1 To Visible_Rows_Count
      'Used to determine next blank line for copying the new validation line.
      Next_Blank_Row = range("A" & Rows.Count).End(xlUp).Row + 1

      range("A" & Next_Blank_Row).Select
      Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
      :=False, Transpose:=False

    Next i


    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' Copies Columns From the Filtered Results Sheet to the New Lines Worksheet
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


    '''''Copies the Alpha_Mon_ID column to the New Lines Sheet'''''''
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


    '''''''Copies the Nomen_Source Column to the New Lines Sheet'''''''

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


    '''''''Copies the Event Code Mapped? Column to the New Lines Sheet'''''''

    'Switches back to Results Sheet to Copy next column
    Sheets("Results").Select
    'Confirms active cell is within the table
    range("A2").Select
    'Selects the first visible cell in column '1'
    ActiveSheet.AutoFilter.range.Offset(1).SpecialCells(xlCellTypeVisible).Cells(1, 1).Select

    range(Selection, Selection.End(xlDown)).Select
    Selection.copy

    Sheets("New Lines").Select

    range("S" & Code_Blank_Line).Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
    :=False, Transpose:=False



    '''''''Copies the Nomenclature Mapped? Column to the New Lines Sheet'''''''

    'Switches back to Results Sheet to Copy next column
    Sheets("Results").Select
    'Confirms active cell is within the table
    range("A2").Select
    'Selects the first visible cell in column '2'
    ActiveSheet.AutoFilter.range.Offset(1).SpecialCells(xlCellTypeVisible).Cells(1, 2).Select

    range(Selection, Selection.End(xlDown)).Select
    Selection.copy

    Sheets("New Lines").Select

    range("T" & Code_Blank_Line).Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
    :=False, Transpose:=False


    'Error handling. No codes found, so skipping
    NoBlanks:
    'MsgBox("No Code for " & EventCode)
      Resume ClearError

    ClearError:
    'Clears variables for next loop
    Visible_Rows_Count = 0

  Next EventCode


  'Re-enables previously disabled settings after all code has run.
  Application.ScreenUpdating = True
  Application.Calculation = xlCalculationAutomatic
  Application.EnableEvents = True

  MsgBox("Program Completed")

End Sub
