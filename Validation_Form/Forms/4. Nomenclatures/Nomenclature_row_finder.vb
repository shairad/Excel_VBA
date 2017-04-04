Private Sub Nomenclature_Row_Finder()

    Dim wb As Workbook
    Dim Table_Obj As ListObject
    Dim Table_ObjIsVisible As Boolean
    Dim Visible_Rows_Count As Integer
    Dim Results_Range As Range
    Dim Val_Vis_Row As Range
    Dim StartCell As Range
    Dim WkNames As Variant
    Dim TblNames As Variant
    Dim DTA_EC_Col As Variant
    Dim Code_ID_Col As Variant
    Dim Mappings_Status_Col As Variant
    Dim ALPHA_NOMEN_ID As Variant


    'DEBUG

    'This disables settings to improve macro performance.
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False


    WkNames = Array("Validated Mappings", "Results", "Validation Sheet")
    TblNames = Array("Mappings_Tbl", "Results_Tbl", "Val_Tbl")


    'Find important column locations
    For i = 0 To UBound(WkNames)

        Sheets(WkNames(i)).Select

        'Finds column locations for columns on the Validated Mappings sheet
        If WkNames(i) = "Validated Mappings" Then
            Range("A1:J1").Select
            Selection.Name = "Header_Row"

            For Each cell In Range("Header_Row")

                If cell = "CODE ID" Then
                    Code_ID_Col = Mid(cell.Address, 2, 1)

                ElseIf cell = "MAPPING STATUS" Then
                    Mappings_Status_Col = Mid(cell.Address, 2, 1)

                End If
            Next cell
        End If

        'Finds column locations for columns on the Results sheet
        If WkNames(i) = "Results" Then

            'finds and stores summary header columns
            Range("A1:M1").Select
            Selection.Name = "Header_Row"

            For Each cell In Range("Header_Row")

                If cell = "DTA_EC" Then
                    DTA_EC_Col = Mid(cell.Address, 2, 1)

                ElseIf cell = "ALPHA_NOMEN_ID" Then
                    ALPHA_NOMEN_ID = Mid(cell.Address, 2, 1)
                End If

            Next cell
        End If

    Next i

    'Converts important sheets to TableStyle
    '''''''''''''''''''''''''''''''''''''''''''''''
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

        Set sht = Worksheets(WkNames(i))    'Sets value
        Set StartCell = Range("A1")    'Start cell used to determine where to begin creating the table range

        'Find Last Row and Column
        LastRow = StartCell.SpecialCells(xlCellTypeLastCell).Row
        LastColumn = StartCell.SpecialCells(xlCellTypeLastCell).Column
        Sheet_Name = WkNames(i)    'Assigns sheet name to a variable as a string

        'Select Range
        sht.Range(StartCell, sht.Cells(LastRow, LastColumn)).Select

        'Creates the table
        Set tbl = ActiveSheet.ListObjects.Add(xlSrcRange, Selection, , xlYes)
        tbl.Name = TblNames(i)    'Names the table
        tbl.TableStyle = "TableStyleLight12"    'Sets table color theme

        Rows("1:1").Select
        With Selection.Font
            .ThemeColor = xlThemeColorDark1
            .TintAndShade = 0
        End With


        'Results Sheet - Adds formulas
        '''''''''''''''''''''''''''''''

        If WkNames(i) = "Results" Then

            'Remove Duplicates from the Results Sheet
            ActiveSheet.Range("Results_Tbl[#All]").RemoveDuplicates Columns:=Array(6, 11, 12), _
                                                                    Header:=xlYes

            'Adds new Mapping note column
            Range("M1").Select
            Selection = "Mapping Note"

            'Eventcode formula
            EventMapped = "=IFERROR(INDEX('Validated Mappings'!" & Mappings_Status_Col & ":" & Mappings_Status_Col & _
                          ",MATCH(" & DTA_EC_Col & "2,'Validated Mappings'!" & Code_ID_Col & ":" & Code_ID_Col & ",0)),0)"

            'Nomenclature formula
            NomenclatureMapped = "=IFERROR(INDEX('Validated Mappings'!" & Mappings_Status_Col & ":" & Mappings_Status_Col & _
                                 ",MATCH(" & ALPHA_NOMEN_ID & "2,'Validated Mappings'!" & Code_ID_Col & ":" & Code_ID_Col & ",0)),0)"

            'Event Code mapped?
            Range("A2").Select
            Selection.Formula = EventMapped

            'Noemcnature mapped?
            Range("B2").Select
            Selection.Formula = NomenclatureMapped

            'Both mapped?
            Range("C2").Select
            Selection.Formula = "=IF(AND(A2 =""Validated"", B2 = ""Validated""),""Both Validated"", 0)"

            'Hides rows which are validated in both columns
            ActiveSheet.ListObjects("Results_Tbl").Range.AutoFilter Field:=3, Criteria1:= _
                                                                    "0"

        End If


        'If the sheet is the validation sheet, thenremove duplicates and create range
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

        If WkNames(i) = "Validation Sheet" Then

            'Removes Duplicates by Event Code from the Validation Sheet
            ActiveSheet.Range("Val_Tbl[#All]").RemoveDuplicates Columns:=9, Header:= _
                                                                xlYes

            'Creates and names the Event_Codes Range which is used in loop
            Range("I2").Select
            Range(Selection, Selection.End(xlDown)).Select
            Selection.Name = "Event_Codes"
        End If

    Next i


    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '       Loops through all event codes to identify lines that need to be handled
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    For Each EventCode In Range("Event_Codes")

        'Filters sheet by event code
        Sheets("Results").Select
        ActiveSheet.ListObjects("Results_Tbl").Range.AutoFilter Field:=6, Criteria1:= _
                                                                EventCode, Operator:=xlAnd

        Set Results_Range = Range("Results_Tbl")

        'Error handling. If no codes are found, then skip the code.
        On Error GoTo NoBlanks

        'Count number of visible rows on the Results sheet
        Visible_Rows_Count = Results_Range.SpecialCells(xlCellTypeVisible).Rows.Count

        'Filters sheet by the current event code
        Sheets("Validation Sheet").Select
        ActiveSheet.ListObjects("Val_Tbl").Range.AutoFilter Field:=9, Criteria1:= _
                                                            EventCode, Operator:=xlAnd

        Set StartCell = Range("A1")

        'finds the first visible row
        Validation_Visible_Row = StartCell.SpecialCells(xlCellTypeLastCell).Row

        'Selects the row
        Rows(Validation_Visible_Row).Select
        Selection.Copy

        Sheets("New Lines").Select

        'Used to determine sheet location when replacing nomenclature values after new lines have been created.
        Code_Blank_Line = Range("A" & Rows.Count).End(xlUp).Row + 1


        '           Creates a new line for each "hit" for a specific code.
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

        For i = 1 To Visible_Rows_Count
            'Used to determine next blank line for copying the new validation line.
            Next_Blank_Row = Range("A" & Rows.Count).End(xlUp).Row + 1

            Range("A" & Next_Blank_Row).Select
            Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                                                                            :=False, Transpose:=False

        Next i


        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        ' Copies Columns From the Filtered Results Sheet to the New Lines Worksheet
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

        '           Copies the Alpha_Mon_ID column to the New Lines Sheet
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

        Sheets("Results").Select

        'Confirms active cell is within the table
        Range("A2").Select

        'Selects the first visible cell in column '12'
        ActiveSheet.AutoFilter.Range.Offset(1).SpecialCells(xlCellTypeVisible).Cells(1, 12).Select

        Range(Selection, Selection.End(xlDown)).Select
        Selection.Copy

        Sheets("New Lines").Select

        Range("M" & Code_Blank_Line).Select
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                                                                        :=False, Transpose:=False


        '           Copies the Nomen_Source Column to the New Lines Sheet
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

        'Switches back to Results Sheet to Copy next column
        Sheets("Results").Select

        'Confirms active cell is within the table
        Range("A2").Select

        'Selects the first visible cell in column '11'
        ActiveSheet.AutoFilter.Range.Offset(1).SpecialCells(xlCellTypeVisible).Cells(1, 11).Select

        Range(Selection, Selection.End(xlDown)).Select
        Selection.Copy

        Sheets("New Lines").Select

        Range("N" & Code_Blank_Line).Select
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                                                                        :=False, Transpose:=False


        '           Copies the Event Code Mapped? Column to the New Lines Sheet
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

        'Switches back to Results Sheet to Copy next column
        Sheets("Results").Select

        'Confirms active cell is within the table
        Range("A2").Select

        'Selects the first visible cell in column '1'
        ActiveSheet.AutoFilter.Range.Offset(1).SpecialCells(xlCellTypeVisible).Cells(1, 1).Select

        Range(Selection, Selection.End(xlDown)).Select
        Selection.Copy

        Sheets("New Lines").Select

        Range("U" & Code_Blank_Line).Select
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                                                                        :=False, Transpose:=False


        '           Copies the Nomenclature Mapped? Column to the New Lines Sheet
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

        'Switches back to Results Sheet to Copy next column
        Sheets("Results").Select

        'Confirms active cell is within the table
        Range("A2").Select

        'Selects the first visible cell in column '2'
        ActiveSheet.AutoFilter.Range.Offset(1).SpecialCells(xlCellTypeVisible).Cells(1, 2).Select

        Range(Selection, Selection.End(xlDown)).Select
        Selection.Copy

        Sheets("New Lines").Select

        Range("V" & Code_Blank_Line).Select
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

    Sheets("New Lines").Select


End Sub
