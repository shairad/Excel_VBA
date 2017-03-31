
    Sub Nomenclature_Notes()

    '
    'This code will take values from a table and put them in an arrao.
    'Then it Will perform changes to the data within the array and then write the array back to the sheet.
    'This changes the values all at once instead of one at a time.
    '
    '

      Dim DataRange As Variant 'Declare array variable
      Dim Irow As Long 'The row variable
      Dim Icol As Integer 'The column variable if you need to loop through multiple columns
      Dim DocType As Variant 'Variable used to store column value
      Dim ControlArray As Variant
      Dim ControlTypeCheck As Variant
      Dim Nomenclature_Val_Check As Variant
      Dim EventCode_Val_Check As Variant
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

        'Creates named Range starting at column E
        Range("E2:T2").Select
        Range(Selection, Selection.End(xlDown)).Select

        Selection.Name = "Data_Range"


      'Array to check DocumentType
      ControlArray = Array("Alpha List", "Alpha Combo", "Discrete Grid", "UltraGrid", "PowerGrid", "Multi")

      'Saves range to array
      DataRange = range("Data_Range").Value 'writes the named data range to the array variable

        For Irow = 1 To UBound(DataRange) 'Loops through all rows within the range.
          DocType = DataRange(Irow, 1)
          ControlTypeCheck = DataRange(Irow, 9)
          Nomenclature_Val_Check = DataRange(Irow, 16)
          EventCode_Val_Check = DataRange(Irow, 15)

          'Checks if control type is within the array.
          IsInArray = Not IsError(Application.Match(ControlTypeCheck, ControlArray, 0))

          If IsInArray = TRUE _
            And Nomenclature_Val_Check = "0" _
            And EventCode_Val_Check = "0" _
            Then

            DataRange(Irow, 13) = "This nomenclature and event code are not mapped and should be if this will be used to complete the measure."
            DataRange(IRow, 14) = "PCST"


          ElseIf IsInArray = TRUE _
            And Nomenclature_Val_Check = "Validated" _
            And EventCode_Val_Check = "0" _
            Then

            DataRange(Irow, 13) = "This nomenclature is mapped but the event code will need to be mapped if this will be used to complete the measure."
            DataRange(IRow, 14) = "PCST"

          ElseIf IsInArray = TRUE _
            And Nomenclature_Val_Check = "0" _
            And EventCode_Val_Check = "Validated" _
            Then

            DataRange(Irow, 13) = "This event code is mapped but the nomenclature is not mapped and should be if this will be used to complete the measure."
            DataRange(IRow, 14) = "Consulting"

          End If

        Next Irow


      'Write the updated DataRange Array to the excel file
      range("Data_Range").Value = DataRange

      're-enables settings previously disabled
      Application.ScreenUpdating = True
      Application.Calculation = xlCalculationAutomatic
      Application.EnableEvents = True

    End Sub
