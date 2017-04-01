Sub Summary_Create_Lookup_Tables()


Dim wb As Workbook
Dim Table_Obj As ListObject
Dim StartCell As Range
Dim WkNames As Variant
Dim TblNames As Variant
Dim PivotNames As Variant
Dim PivotSheetNames As Variant
Dim lastrow As Long
Dim LastColumn As Long
Dim rList As Range


'DEBUG

'This disables settings to improve macro performance.
'Application.ScreenUpdating = False
'Application.Calculation = xlCalculationManual
'Application.EnableEvents = False


WkNames = Array("Potential Mapping Issues", "Unmapped Codes", "Clinical Documentation")
TblNames = Array("Potential_Table", "Unmapped_Table", "Clinical_Table")
PivotNames = Array("Potential_Pivot", "Unmapped_Pivot", "Clinical_Pivot")
PivotSheetNames = Array("Potential_Summary_Pivot", "Unmapped_Summary_Pivot", "Clinical_Summary_Pivot")


For i = 0 To UBound(WkNames)

  CurrentWkName = WkNames(i)
  CurrentTblName = TblNames(i)
  CurrentPivotName = PivotNames(i)
  CurrentPivotSheetName = PivotSheetNames(i)

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
  Set StartCell = Range("A2") 'Start cell used to determine where to begin creating the table range

'Find Last Row and Column
  lastrow = StartCell.SpecialCells(xlCellTypeLastCell).Row
  LastColumn = StartCell.SpecialCells(xlCellTypeLastCell).Column
  Sheet_Name = WkNames(i) 'Assigns sheet name to a variable as a string

'Select Range
  sht.Range(StartCell, sht.Cells(lastrow, LastColumn)).Select

'Creates the table
  Set tbl = ActiveSheet.ListObjects.Add(xlSrcRange, Selection, , xlYes)
  tbl.Name = TblNames(i) 'Names the table
  tbl.TableStyle = "TableStyleLight12" 'Sets table color theme

  Rows("2:2").Select
  With Selection.Font
    .ThemeColor = xlThemeColorDark1
    .TintAndShade = 0
  End With


  'Creates a new sheet which will house the validated codes pivot table
    With ThisWorkbook
        .Sheets.Add(After:=.Sheets(.Sheets.Count)).Name = CurrentPivotSheetName
    End With

    Sheets(CurrentWkName).Select
    Range("B2").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
      CurrentTblName, Version:=6).CreatePivotTable TableDestination:= _
    CurrentPivotSheetName & "!R1C1", TableName:=CurrentPivotName, DefaultVersion:=6

    Sheets(CurrentPivotSheetName).Select
    Cells(1, 1).Select


    ActiveSheet.PivotTables(CurrentPivotName).AddDataField ActiveSheet.PivotTables( _
    CurrentPivotName).PivotFields("Source"), "Count of Source", xlCount

    With ActiveSheet.PivotTables(CurrentPivotName).PivotFields("Registry")
        .Orientation = xlRowField
        .Position = 1
    End With


    With ActiveSheet.PivotTables(CurrentPivotName).PivotFields("Measure")
        .Orientation = xlRowField
        .Position = 2
    End With

  'Sets pivot table layout to OUTLINE
    ActiveSheet.PivotTables(CurrentPivotName).RowAxisLayout xlOutlineRow


  'Turns on repeat blank lines
    ActiveSheet.PivotTables(CurrentPivotName).RepeatAllLabels xlRepeatLabels

  'Sets empty values to 0 which helps in a couple places! but also allows the below autofill to have a range reference'
    ActiveSheet.PivotTables(CurrentPivotName).NullString = "0"


    Range("D1").Select

    lastrow = ActiveSheet.Range("C2").End(xlDown).Row

    Sheets(CurrentPivotSheetName).Select
    Range("D2").Select
    ActiveCell.Formula = "=IF(B2 <>"""",CONCATENATE(A2,""|"",B2),"""")"


    With ActiveSheet.Range("D2")
        .AutoFill Destination:=Range("D2:D" & lastrow&)
    End With

Next i


End Sub
