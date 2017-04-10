Sub BORIS_PCST()

'
'
' PURPOSE: Creates all the PCST files formatted properly.
'	USE:	Open validation form which will be used to make PCST files. Run program and follow on screen prompts.
' AUTHOR: Jonathan Adams
'
'
'

    Dim wb As Workbook
    Dim FirstWkbk As Workbook
    Dim sht As Worksheet
    Dim Sheet As Worksheet
    Dim Table_Obj As ListObject
    Dim tbl As ListObject
    Dim Validation_File_Name As Variant
    Dim cValue As Variant
    Dim StartCell As Range
    Dim rList As Range
    Dim User_Name As String
    Dim Project_Name As String
    Dim Save_Path As String
    Dim Code_Sheet As String
    Dim Current_Value As String
    Dim Sheet_Name As String
    Dim Name_Input_Checker As Integer
    Dim Confirm_Scrubbed As Integer
    Dim Folder_Check As Integer
    Dim Project_Name_Checker As Integer
    Dim LastRow As Long
    Dim LastColumn As Long
    Dim Next_Blank_Row As Long
    Dim Table_ObjIsVisible As Boolean
    Dim Checker_Health_Maint As Boolean
    Dim Val_Wk_Array As Variant
    Dim Val_Tbl_Name_Array As Variant
    Dim CurrentSheet As String


    ' 'This disables settings to improve macro performance.
    ' Application.ScreenUpdating = False
    ' Application.Calculation = xlCalculationManual
    ' Application.EnableEvents = False


    ' DEBUG
    User_Name = "ja052464"
    Project_Name = "Test"
    Validation_File_Name = ActiveWorkbook.Name

' TODO Look at optimizing the unique code source code. Beginning on line 329

    Val_Wk_Array = Array("Clinical Documentation", "Unmapped Codes", "Health Maintenance Summary")
    Val_Tbl_Name_Array = Array("Clinical_Table", "Unmapped_Table", "Health_Maint_Table")


    '''''''''''FORMATS THE WORKSHEETS FOR COPYING TO THE NEW WORKBOOK'''''''''''''

    For i = 0 to UBound(Val_Wk_Array)

        Sheets(Val_Wk_Array(i)).Select

        'table can not be created if autofilters are on
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

        'Sets all cells on sheet to a table.
        Set sht = Worksheets(Val_Wk_Array(i))

        ' Health Maint Sheet data starts on a different row
        If Val_Wk_Array(i) <> "Health Maintenance Summary" Then
            Set StartCell = Range("A2")
            LastRow = StartCell.SpecialCells(xlCellTypeLastCell).Row
            LastColumn = StartCell.SpecialCells(xlCellTypeLastCell).Column

            'Select Range
            sht.Range(StartCell, sht.Cells(LastRow, LastColumn)).Select
        Else
          Range("A5").Select
          Range(Selection, Selection.End(xlDown)).Select
          Range(Selection, Selection.End(xlToRight)).Select

        End if


        'Converts range to table.
        Set tbl = ActiveSheet.ListObjects.Add(xlSrcRange, Selection, , xlYes)
        tbl.Name = Val_Tbl_Name_Array(i)
        tbl.TableStyle = "TableStyleLight12"

        'Filters to remove blank lines
        ActiveSheet.ListObjects(1).Range.AutoFilter Field:=5, _
        Criteria1:="<>"

    Next i


    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'CREATES THE SOURCE CODE SHEET AND TABLE FOR LOOP ON THE VALIDATION FORM
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    'Disables screen alert which would prompt user to confirm sheet deletion.
    Application.DisplayAlerts = False

    'Checks to see if sources list sheet already exists and if so deletes the worksheet so a new one can be created.
    For Each Sheet In Worksheets
        If Sheet.Name = "Sources List" Then
            exists = True
            Sheet.Delete
        End If
    Next Sheet

    're-enables screen alert after handling source code sheet deletion.
    Application.DisplayAlerts = True

    'Creates a new sheet titled 'Sources List'
    With ThisWorkbook
        .Sheets.Add(After:=.Sheets(.Sheets.Count)).Name = "Sources List"
        Range("A1").Value = "Sources"
    End With

    ' Assigns starting value to Next Blank Row
    Next_Blank_Row = Sheets("Sources List").Range("A" & Rows.Count).End(xlUp).Row + 1
    ' Loops through the sheets to find all the sources and put them on the sources list sheet
    For i = 0 to UBound(Val_Wk_Array)
        CurrentSheet = Val_Wk_Array(i)
        Sheets(CurrentSheet).Select

        If CurrentSheet <> "Health Maintenance Summary" Then
            'Copies sources from the sheets to the sources list
            Range(ActiveSheet.Range("E2"), ActiveSheet.Range("E2").End(xlDown)).Copy Sheets("Sources List").Range("A" & Next_Blank_Row)
            Sheets("Sources List").Rows(Next_Blank_Row & ":" & Next_Blank_Row).Delete Shift:=xlUp
        Else
            'Copies sources from the sheets to the sources list
            Range(ActiveSheet.Range("K5"), ActiveSheet.Range("K5").End(xlDown)).Copy Sheets("Sources List").Range("A" & Next_Blank_Row)
            Sheets("Sources List").Rows(Next_Blank_Row & ":" & Next_Blank_Row).Delete Shift:=xlUp
        End if
        'Finds next blank row to add additional sources.
        Next_Blank_Row = Sheets("Sources List").Range("A" & Rows.Count).End(xlUp).Row + 1
    Next i


    'Create named range of the sources
    Sheets("Sources List").Select
    Range("A1").Select
    Range(Selection, Selection.End(xlDown)).Select

    ' formats selected as a table
    Set tbl = ActiveSheet.ListObjects.Add(xlSrcRange, Selection, , xlYes)
    tbl.Name = "Sources_Table"
    tbl.TableStyle = "TableStyleLight12"

    ' Removes Duplicates
    ActiveSheet.Range("Sources_Table[#All]").RemoveDuplicates Columns:=1, Header _
            :=xlYes

    ' Names the range
    Range("A2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Name = "Sources_List"



    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' PERFORMS CHECKS ON THE VALIDATION FORM TO CONFIRM FORMAT IS CORRECT BEFORE PROCEEDING ANY FURTHER
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    ' Checks the Health Maintenance Summary
    If LCase(Sheets("Health Maintenance Summary").Range("K5").value) <> "source" Then

        MsgBox("Program has detected a possible error with the Validation Form layout" & vbNewLine & vbNewLine & _
        "Program expected Column K on the Health Maintenance Summary sheet to be 'Source'. Please resolve the issue and then run again.")

        'Re-enables previously disabled settings after all code has run.
        Application.ScreenUpdating = True
        Application.Calculation = xlCalculationAutomatic
        Application.EnableEvents = True
        Exit Sub
    End If


    '
    ''''''''''''''''''''''''''''''''''''''''
    '         CREATE NEW WORKBOOK
    ''''''''''''''''''''''''''''''''''''''''
    '


    'Loop through the sources
    For Each Source_Name In Range("Sources_List")
        Set wb = Workbooks.Add    'Opens a new workbook

        'Saves the new workbook
        With NewBook
            ChDir "C:\Users\" & User_Name & "\Documents\" & Project_Name & "_" & "PCST_Files"
            ActiveWorkbook.SaveAs Filename:= _
                    "C:\Users\" & User_Name & "\Documents\" & Project_Name & "_" & "PCST_Files\" & Source_Name, FileFormat:= _
                    xlOpenXMLWorkbookMacroEnabled, CreateBackup:=False
        End With

        'Selects new workbook
        Windows(Source_Name & ".xlsm").Activate

        'Populates basic sheets on new workbook
        With ActiveWorkbook
            .Sheets.Add(After:=.Sheets(.Sheets.Count)).Name = "Unmapped Codes"
            .Sheets.Add(After:=.Sheets(.Sheets.Count)).Name = "Clinical Documentation"
            .Sheets.Add(After:=.Sheets(.Sheets.Count)).Name = "Health Maintenance Summary"
            .Sheets.Add(After:=.Sheets(.Sheets.Count)).Name = "Source_Code_Systems"
        End With




        For i = 0 to UBound(Val_Wk_Array)
            CurrentSheet = Val_Wk_Array(i)
            CurrentTable = Val_Tbl_Name_Array(i)

            Windows(Validation_File_Name).Activate
            Sheets(CurrentSheet).Select

          If CurrentSheet <> Val_Wk_Array(2) Then
            ' Finds the location of the Registry Column
              Range("A2:K2").Select
              Selection.Name = "Header_row"

              For each cell in Range("Header_row")
                If cell = "Registry" Then
                  start_cell = cell.Address
                  exit for
                End if
              Next cell

          Else
          ' Finds the location of the Registry Column
            Range("A5:K5").Select
            Selection.Name = "Header_row"

            For each cell in Range("Header_row")
              If cell = "EXPECT_NAME" Then
                start_cell = cell.Address
                exit for
              End if
            Next cell

          end if


          Sheets(CurrentSheet).Select
          ActiveSheet.ListObjects(1).Range.AutoFilter Field:=5, _
                  Criteria1:="<>"
              Range(start_cell).Select
              Range(Selection, Selection.End(xlDown)).Select
              Range(Selection, Selection.End(xlToRight)).Select
              Selection.Copy

          'Selects the newly created excel file and pastes copied cells onto unmapped codes sheet
          Windows(Source_Name & ".xlsm").Activate
          Sheets(CurrentSheet).Select
          Range("A1").Select
          Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                  :=False, Transpose:=False
        Next i

        ' Formats the new sheets as tables
        For i = 0 to UBound(Val_Wk_Array)
            'Confirms Selection of new workbook
            Windows(Source_Name & ".xlsm").Activate
            Sheets(Val_Wk_Array(i)).Select

            'Sets all cells on sheet to a table.
            Set sht = Worksheets(Val_Wk_Array(i))
            Set StartCell = Range("A1")

            'Find Last Row and Column
            LastRow = StartCell.SpecialCells(xlCellTypeLastCell).Row
            LastColumn = StartCell.SpecialCells(xlCellTypeLastCell).Column

            'Select Range
            sht.Range(StartCell, sht.Cells(LastRow, LastColumn)).Select

            'Converts range to table.
            Set tbl = ActiveSheet.ListObjects.Add(xlSrcRange, Selection, , xlYes)
            tbl.Name = Val_Tbl_Name_Array(i)
            tbl.TableStyle = "TableStyleLight12"

            'Filters to remove blank lines
            ActiveSheet.ListObjects(1).Range.AutoFilter Field:=5, _
            Criteria1:="<>"

        Next i


        ' Formats the unmapped code sheet for the code ID's
        Sheets(Val_Wk_Array(1)).Select

        'Removes duplicates if any exist by Raw Code and Raw Display
        ActiveSheet.Range(Val_Tbl_Name_Array(1)&"[#All]").RemoveDuplicates Columns:=Array(6, 7, 8 _
            ), Header:=xlYes


        ' Sets header for code short name column
        Range("K1").Value = "Code Short Name"

        'Copies code id to short code id column
        Range("F2").Select
        Range(Selection, Selection.End(xlDown)).Select
        Selection.Copy
        Range("K2").Select
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                :=False, Transpose:=False

        ' Clears any autofilters
        ActiveSheet.autofilter.showalldata

        ' filters by concept -> registry for easy reviewing in final product
        With ActiveWorkbook.ActiveSheet.ListObjects(1).Sort
            .SortFields.Add Key:=Range(Val_Tbl_Name_Array(1)&"[Concept]"), SortOn:= _
                xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
            .SortFields.Add Key:=Range(Val_Tbl_Name_Array(1)&"[Measure]"), SortOn:= _
                xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
            .SortFields.Add Key:=Range(Val_Tbl_Name_Array(1)&"[Registry]"), SortOn:= _
                xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
            .Apply
        End With


Next Source_Name

' TODO Optimize code with loops!

        '''''''''''''CREATES THE SOURCE CODE SHEET'''''''''''''

        ' Selects the unmapped codes sheet and copies data
        Sheets(Val_Wk_Array(1)).Select
        ActiveSheet.ListObjects(1).Range.AutoFilter Field:=5, _
                Criteria1:=Source_Name, Operator:=xlAnd

        Range("K1").Select
        Range(Selection, Selection.End(xlDown)).Select
        Selection.Copy

        'Pastes unmapped codes list onto the source code systems sheet
        Sheets("Source_Code_Systems").Select
        Range("A1").Select
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                :=False, Transpose:=False


        ' Formats Source_Code_Systems Sheet
        Set sht = Worksheets("Source_Code_Systems")
        Set StartCell = Range("A1")

        'Refresh UsedRange
        Worksheets("Source_Code_Systems").UsedRange

        'Find Last Row and Column
        LastRow = StartCell.SpecialCells(xlCellTypeLastCell).Row
        LastColumn = StartCell.SpecialCells(xlCellTypeLastCell).Column

        'Select Range
        sht.Range(StartCell, sht.Cells(LastRow, LastColumn)).Select

        Set tbl = ActiveSheet.ListObjects.Add(xlSrcRange, Selection, , xlYes)
        tbl.Name = "Code_ID_Table"
        tbl.TableStyle = "TableStyleLight9"

        Range("Code_ID_Table[[#Headers],[Code Short Name]]").Select
        Application.CutCopyMode = False
        ActiveSheet.Range("Code_ID_Table[#All]").RemoveDuplicates Columns:=1, Header:= _
                xlYes



        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        ' Checks to determine how many unique code ID's there are for this source.
        ' If there is only 1 source "72" then set range to one cell, otherwise select all cells and set the range.
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

        'If there are no unmapped codes for this source, then set code source to 72 only.
        Range("A2").Select
        If Selection.Value = "" Then
            Selection.Value = "72"
        End If
        ' If A3 is empty and thus only 1 code id, then set the range of Code_ID_List to just cell A2
        Range("A3").Select
        If Selection.Value = "" Then
            Range("A2").Select
            Selection.Name = "Code_ID_List"
        Else
            Range("A2").Select
            Range(Selection, Selection.End(xlDown)).Select
            Selection.Name = "Code_ID_List"
        End If


        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        '             Creates a new sheet and names the sheet with the current source
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

        'Loops through each unique Code ID for this source and creates a sheet with the relavant data.
        For Each code In Range("Code_ID_List")

            With ActiveWorkbook
                .Sheets.Add(After:=.Sheets(.Sheets.Count)).Name = code
            End With

            Code_Sheet = code

' TODO Create a range and use a loop and an arra to populate sheet headers

            'Special instructions for code set 72
            If code = "72" Then

                Sheets(Code_Sheet).Select

                'Populates sheet 72 headers
                Range("A1").Select
                ActiveCell.FormulaR1C1 = "Registry"
                Range("B1").Select
                ActiveCell.FormulaR1C1 = "Measure"
                Range("C1").Select
                ActiveCell.FormulaR1C1 = "Concept"
                Range("D1").Select
                ActiveCell.FormulaR1C1 = "Source"
                Range("E1").Select
                ActiveCell.FormulaR1C1 = "DocumentType"
                Range("F1").Select
                ActiveCell.FormulaR1C1 = "Name"
                Range("G1").Select
                ActiveCell.FormulaR1C1 = "Section"
                Range("H1").Select
                ActiveCell.FormulaR1C1 = "DTA"
                Range("I1").Select
                ActiveCell.FormulaR1C1 = "EventCode"
                Range("J1").Select
                ActiveCell.FormulaR1C1 = "EventDisplay"
                Range("K1").Select
                ActiveCell.FormulaR1C1 = "ESH"
                Range("L1").Select
                ActiveCell.FormulaR1C1 = "ControlType"
                Range("M1").Select
                ActiveCell.FormulaR1C1 = "NomenclatureID"
                Range("N1").Select
                ActiveCell.FormulaR1C1 = "Nomenclature"
                Range("O1").Select
                ActiveCell.FormulaR1C1 = "TaskAssay"
                Range("P1").Select
                ActiveCell.FormulaR1C1 = "Notes"
                Range("Q1").Select
                ActiveCell.FormulaR1C1 = "Comments"
                Range("R1").Select
                ActiveCell.FormulaR1C1 = "Standard Code"
                Range("S1").Select
                ActiveCell.FormulaR1C1 = "Standard Coding System"
                Range("Q2").Select

                ''''''''''''Copies Clinical Documentation to 72'''''''''''''
' TODO Look to see what we can do to clean this up
                Sheets("Clinical Documentation").Select

                'Filters for only the current source
                ActiveSheet.ListObjects("Clinical_Table").Range.AutoFilter Field:=5, _
                        Criteria1:=Source_Name, Operator:=xlAnd

                Set Table_Obj = ActiveSheet.ListObjects(1)

                'Checks filtered table for visible data
                If Table_Obj.Range.SpecialCells(xlCellTypeVisible).Areas.Count > 1 Then
                    Table_ObjIsVisible = True
                Else
                    Table_ObjIsVisible = False
                End If

                'If data is visible, then copy visible data
                If Table_ObjIsVisible = True Then

                    Sheets("Clinical Documentation").Select
                    Columns("B:O").Select
                    Selection.Copy
                    Sheets(Code_Sheet).Select
                    Range("A1").Select
                    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                            :=False, Transpose:=False

                    Sheets("Clinical Documentation").Select
                    Columns("Q:Q").Select
                    Application.CutCopyMode = False
                    Selection.Copy
                    Sheets(Code_Sheet).Select
                    Columns("O:O").Select
                    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                            :=False, Transpose:=False

                    Sheets("Clinical Documentation").Select
                    Columns("R:R").Select
                    Application.CutCopyMode = False
                    Selection.Copy
                    Sheets(Code_Sheet).Select
                    Range("P1").Select
                    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                            :=False, Transpose:=False
                End If


                ''''''''''Copies unmapped codes to 72 sheet''''''''''
' TODO clean up
                Sheets("Unmapped Codes").Select

                'Applies filters for only this source and code being currently reviewed.
                ActiveSheet.ListObjects("Unmapped_Table").Range.AutoFilter Field:=5, _
                        Criteria1:=Source_Name, Operator:=xlAnd
                ActiveSheet.ListObjects("Unmapped_Table").Range.AutoFilter Field:=12, _
                        Criteria1:=code, Operator:=xlAnd

                Set Table_Obj = ActiveSheet.ListObjects(1)

                'Checks table for visible data
                If Table_Obj.Range.SpecialCells(xlCellTypeVisible).Areas.Count > 1 Then
                    Table_ObjIsVisible = True
                Else
                    Table_ObjIsVisible = False
                End If

                'If data is visible then copy data
                If Table_ObjIsVisible = True Then

                    Range("B2:F2").Select
                    Range(Selection, Selection.End(xlToLeft)).Select
                    Range("B2:F2").Select
                    Range(Selection, Selection.End(xlDown)).Select
                    Selection.Copy

                    Sheets(Code_Sheet).Select

                    Next_Blank_Row = Range("A" & Rows.Count).End(xlUp).Row + 1
                    'Selects next blank row
                    Range("A" & Next_Blank_Row).Select
                    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                            :=False, Transpose:=False

                    Sheets("Unmapped Codes").Select

                    Range("G2:I2").Select
                    Range(Selection, Selection.End(xlDown)).Select
                    Application.CutCopyMode = False
                    Selection.Copy

                    Sheets(Code_Sheet).Select
                    Range("I" & Next_Blank_Row).Select
                    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                            :=False, Transpose:=False
                End If


                '''''''''Populates Health Maintenance to CS 72''''''''''''''
' TODO Cleanup
                Sheets("Health Maintenance Summary").Select

                ActiveSheet.ListObjects("Health_Maint_Table").Range.AutoFilter Field:=11, _
                        Criteria1:=Source_Name, Operator:=xlAnd

                Set tbl = ActiveSheet.ListObjects(1)

                'Checks table for visible data
                If tbl.Range.SpecialCells(xlCellTypeVisible).Areas.Count > 1 Then
                    Table_ObjIsVisible = True
                Else:
                    Table_ObjIsVisible = tbl.Range.SpecialCells(xlCellTypeVisible).Rows.Count > 1
                End If

                'If data is visible then copy visible data
                If Table_ObjIsVisible = True Then

                    'Copies Sources Column
                    Range("K2:K" & Cells.SpecialCells(xlCellTypeLastCell).Row).Select
                    Selection.Copy
                    Sheets(Code_Sheet).Select

                    'Uses The CodeSet 72 column "Sources" to determine next blank row
                    Next_Blank_Row = Range("D" & Rows.Count).End(xlUp).Row + 1

                    'Pastes Sources on new sheet
                    Range("D" & Next_Blank_Row).Select
                    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                            :=False, Transpose:=False

                    'Copies Expect_Meaning Column
                    Sheets("Health Maintenance Summary").Select
                    Range("B2:B" & Cells.SpecialCells(xlCellTypeLastCell).Row).Select
                    Selection.Copy
                    Sheets(Code_Sheet).Select

                    'Pastes Expect_Meaning on new sheet
                    Range("F" & Next_Blank_Row).Select
                    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                            :=False, Transpose:=False

                    'Copies Satisfier_Meaning Column
                    Sheets("Health Maintenance Summary").Select
                    Range("G2:G" & Cells.SpecialCells(xlCellTypeLastCell).Row).Select
                    Selection.Copy
                    Sheets(Code_Sheet).Select

                    'Pastes Satisfier_Meaning on new sheet
                    Range("G" & Next_Blank_Row).Select
                    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                            :=False, Transpose:=False

                    'Copies Entry_Type Column
                    Sheets("Health Maintenance Summary").Select
                    Range("C2:C" & Cells.SpecialCells(xlCellTypeLastCell).Row).Select
                    Selection.Copy
                    Sheets(Code_Sheet).Select

                    'Pastes Entry_Type on new sheet
                    Range("L" & Next_Blank_Row).Select
                    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                            :=False, Transpose:=False

                    'Copies Event_CD Column
                    Sheets("Health Maintenance Summary").Select
                    Range("I2:I" & Cells.SpecialCells(xlCellTypeLastCell).Row).Select
                    Selection.Copy
                    Sheets(Code_Sheet).Select

                    'Pastes Event_CD on new sheet
                    Range("I" & Next_Blank_Row).Select
                    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                            :=False, Transpose:=False

                    'Copies Event_CD_DISP Column
                    Sheets("Health Maintenance Summary").Select
                    Range("J2:J" & Cells.SpecialCells(xlCellTypeLastCell).Row).Select
                    Selection.Copy
                    Sheets(Code_Sheet).Select

                    'Pastes Event_CD_DISP on new sheet
                    Range("J" & Next_Blank_Row).Select
                    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                            :=False, Transpose:=False
                End If


                '''''''''Populates headers for all other sheets''''''''''
' TODO use loop and array to populate headers
            Else
                Sheets(Code_Sheet).Select

                Range("A1").Select
                ActiveCell.FormulaR1C1 = "Registry"
                Range("B1").Select
                ActiveCell.FormulaR1C1 = "Measure"
                Range("C1").Select
                ActiveCell.FormulaR1C1 = "Concept"
                Range("D1").Select
                ActiveCell.FormulaR1C1 = "Source"
                Range("E1").Select
                ActiveCell.FormulaR1C1 = "DocumentType"
                Range("F1").Select
                ActiveCell.FormulaR1C1 = "Name"
                Range("G1").Select
                ActiveCell.FormulaR1C1 = "Section"
                Range("H1").Select
                ActiveCell.FormulaR1C1 = "DTA"
                Range("I1").Select
                ActiveCell.FormulaR1C1 = "Code"
                Range("J1").Select
                ActiveCell.FormulaR1C1 = "Display"
                Range("K1").Select
                ActiveCell.FormulaR1C1 = "ESH"
                Range("L1").Select
                ActiveCell.FormulaR1C1 = "ControlType"
                Range("M1").Select
                ActiveCell.FormulaR1C1 = "NomenclatureID"
                Range("N1").Select
                ActiveCell.FormulaR1C1 = "Nomenclature"
                Range("O1").Select
                ActiveCell.FormulaR1C1 = "vlookup"
                Range("P1").Select
                ActiveCell.FormulaR1C1 = "Team"
                Range("Q1").Select
                ActiveCell.FormulaR1C1 = "Comments"
                Range("R1").Select
                ActiveCell.FormulaR1C1 = "Standard Code"
                Range("S1").Select
                ActiveCell.FormulaR1C1 = "Standard Coding System"
                Range("Q2").Select

                'Filters unmapped codes table for current source and code within loop.
                Sheets("Unmapped Codes").Select
                ActiveSheet.ListObjects("Unmapped_Table").Range.AutoFilter Field:=5, _
                        Criteria1:=Source_Name, Operator:=xlAnd
                ActiveSheet.ListObjects("Unmapped_Table").Range.AutoFilter Field:=12, _
                        Criteria1:=code, Operator:=xlAnd

                'Sets variable to the table on the active sheet.
                Set Table_Obj = ActiveSheet.ListObjects(1)

                'Checks filtered table for visible data.
                If Table_Obj.Range.SpecialCells(xlCellTypeVisible).Areas.Count > 1 Then
                    Table_ObjIsVisible = True
                Else
                    Table_ObjIsVisible = False
                End If

                'If data is visible then copy data.
                If Table_ObjIsVisible = True Then

                    Sheets("Unmapped Codes").Select
                    Range("B2:F2").Select
                    Range(Selection, Selection.End(xlDown)).Select
                    Selection.Copy

                    Sheets(Code_Sheet).Select

                    'Finds next blank row
                    Next_Blank_Row = Range("A" & Rows.Count).End(xlUp).Row + 1

                    'Pastes data on next blank row
                    Range("A" & Next_Blank_Row).Select
                    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                            :=False, Transpose:=False

                    Sheets("Unmapped Codes").Select

                    Range("G2:I2").Select
                    Range(Selection, Selection.End(xlDown)).Select
                    Application.CutCopyMode = False
                    Selection.Copy

                    Sheets(Code_Sheet).Select
                    Range("I" & Next_Blank_Row).Select
                    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                            :=False, Transpose:=False

                End If

            End If

        Next code


        '''''''POPULATES NOMENCLATURE - PATIENT CARE SHEET IF NEEDED''''''''


        Sheets("Clinical Documentation").Select

        'Filters table for current source.
        ActiveSheet.ListObjects("Clinical_Table").Range.AutoFilter Field:=5, _
                Criteria1:=Source_Name, Operator:=xlAnd
        ActiveSheet.ListObjects("Clinical_Table").Range.AutoFilter Field:=15, _
                Criteria1:="<>"

        'Eventually update this to filter out all rows which ARE MAPPED CORRECTLY. To only leave
        'incorrect rows.
        'ActiveSheet.ListObjects("Clinical_Table").Range.AutoFilter Field:=18, _
         'Criteria1:="<>"

        Set Table_Obj = ActiveSheet.ListObjects(1)

        'Checks filtered table for visible data.
        If Table_Obj.Range.SpecialCells(xlCellTypeVisible).Areas.Count > 1 Then
            Table_ObjIsVisible = True
        Else
            Table_ObjIsVisible = False
        End If

        'If data is visible then copy data.
        If Table_ObjIsVisible = True Then
            'Check to see if Nomenclature - Patient Care sheet already exists
            For Each Sheet In Worksheets
                If Sheet.Name = "Nomenclature - Patient Care" Then
                    exists = True
                    Exit For
                Else
                    exists = False
                End If
            Next Sheet

            'If sheet does NOT exist, then create the sheet
            If exists = False Then
                ActiveWorkbook.Sheets.Add(After:=Worksheets(1)).Name = "Nomenclature - Patient Care"

                'Populate Headers
                Sheets("Nomenclature - Patient Care").Select
                Range("A1").Select
                ActiveCell.FormulaR1C1 = "Registry"
                Range("B1").Select
                ActiveCell.FormulaR1C1 = "Measure"
                Range("C1").Select
                ActiveCell.FormulaR1C1 = "Concept"
                Range("D1").Select
                ActiveCell.FormulaR1C1 = "Source"
                Range("E1").Select
                ActiveCell.FormulaR1C1 = "DocumentType"
                Range("F1").Select
                ActiveCell.FormulaR1C1 = "Name"
                Range("G1").Select
                ActiveCell.FormulaR1C1 = "Section"
                Range("H1").Select
                ActiveCell.FormulaR1C1 = "DTA"
                Range("I1").Select
                ActiveCell.FormulaR1C1 = "Code"
                Range("J1").Select
                ActiveCell.FormulaR1C1 = "Display"
                Range("K1").Select
                ActiveCell.FormulaR1C1 = "ESH"
                Range("L1").Select
                ActiveCell.FormulaR1C1 = "ControlType"
                Range("M1").Select
                ActiveCell.FormulaR1C1 = "NomenclatureID"
                Range("N1").Select
                ActiveCell.FormulaR1C1 = "Nomenclature"
                Range("O1").Select
                ActiveCell.FormulaR1C1 = "vlookup"
                Range("P1").Select
                ActiveCell.FormulaR1C1 = "Team"
                Range("Q1").Select
                ActiveCell.FormulaR1C1 = "Comments"
                Range("R1").Select
                ActiveCell.FormulaR1C1 = "Standard Code"
                Range("S1").Select
                ActiveCell.FormulaR1C1 = "Standard Coding System"
                Range("Q2").Select
            End If

            'Populates the Nomenclature - Patient Care Sheet with data from Clinical Documentation
            Sheets("Clinical Documentation").Select
            Range("B2:Q2").Select
            Range(Selection, Selection.End(xlDown)).Select
            Selection.Copy

            Sheets("Nomenclature - Patient Care").Select

            'Selects next blank row
            Next_Blank_Row = Range("A" & Rows.Count).End(xlUp).Row + 1
            Range("A" & Next_Blank_Row).Select
            Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                    :=False, Transpose:=False

            Set StartCell = Range("A1")
            Set sht = Worksheets("Nomenclature - Patient Care")

            'Finds last row with text
            LastRow = StartCell.SpecialCells(xlCellTypeLastCell).Row

            'Removes duplicates from sheet by source, nom ID and Nom description
            ActiveSheet.Range("$A$1:$O$" & LastRow).RemoveDuplicates Columns:=Array(4, 14, 15), _
                    Header:=xlYes

        Else
            'Do Nothing
        End If


        '''''''''''Loops through sheets and delets unneeded sheets''''''''''''

        'Disables on screen confirm prompt
        Application.DisplayAlerts = False

        For Each Sheet In Worksheets
            If Sheet.Name = "Unmapped Codes" _
                    Or Sheet.Name = "Health Maintenance Summary" _
                    Or Sheet.Name = "Clinical Documentation" _
                    Or Sheet.Name = "Source_Code_Systems" _
                    Or Sheet.Name = "Sheet1" _
                    Then
                Sheet.Delete
            End If
        Next Sheet

        're-enables on screen prompt
        Application.DisplayAlerts = True

        '''''''Creates Index Sheet'''''''''

        'Creates index sheet
        ActiveWorkbook.Sheets.Add(Before:=Worksheets(1)).Name = "Index Sheet"

        Range("A1").Select
        Selection.Value = "Index Sheet"
        ActiveCell.Offset(1, 0).Select    'Moves down a row

        For Each Sheet In Worksheets
            If Sheet.Name <> "Index Sheet" Then
                ActiveSheet.Hyperlinks.Add Anchor:=Selection, Address:="", SubAddress:="'" & Sheet.Name & "'" & "!A1", TextToDisplay:=Sheet.Name
                ActiveCell.Offset(1, 0).Select    'Moves down a row
            End If
        Next Sheet

        ''''''''Loops through sheets and formats as table'''''''

        For Each Sheet In Worksheets
            'Activates current sheet
            Sheet.Activate

            Set sht = Sheet    'Sets value
            Set StartCell = Range("A1")    'Start cell used to determine where to begin creating the table range

            'Find Last Row and Column
            LastRow = StartCell.SpecialCells(xlCellTypeLastCell).Row
            LastColumn = StartCell.SpecialCells(xlCellTypeLastCell).Column
            Sheet_Name = Sheet.Name    'Assigns sheet name to a variable as a string

            'Select Range
            sht.Range(StartCell, sht.Cells(LastRow, LastColumn)).Select

            'Creates the table
            Set tbl = ActiveSheet.ListObjects.Add(xlSrcRange, Selection, , xlYes)
            tbl.Name = Sheet_Name    'Names the table
            tbl.TableStyle = "TableStyleLight9"    'Sets table color theme
            Columns.AutoFit    'Autofits columns on sheet
            Range("A1").Select    'Selects Cell A1 on sheet. Completely cosmetic.
        Next Sheet

        ''''''''Aligns index sheet'''''''''

        Sheets("INDEX SHEET").Select
        Range("A2").Select
        Range(Selection, Selection.End(xlDown)).Select
        With Selection
            .HorizontalAlignment = xlLeft
            .VerticalAlignment = xlBottom
            .WrapText = False
            .Orientation = 0
            .AddIndent = False
            .IndentLevel = 0
            .ShrinkToFit = False
            .ReadingOrder = xlContext
            .MergeCells = False
        End With


        '''''''Saves the new workbook'''''''''

        Workbooks(Source_Name & ".xlsm").Close SaveChanges:=True
        Windows(Validation_File_Name).Activate    'Switches back to old workbook to begin next loop

    Next Source_Name    'Start over with a new source

    'Re-enables previously disabled settings after all code has run.
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True

    'Notifies user that the program has completed.
    MsgBox ("Your PCST Files have been created. Folder is loctated within your My Documents.")

    End_Program:
        'Re-enables previously disabled settings after all code has run.
        Application.ScreenUpdating = True
        Application.Calculation = xlCalculationAutomatic
        Application.EnableEvents = True


End Sub
