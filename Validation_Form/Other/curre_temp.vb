Sub BORIS_PCST()

'
'
' PURPOSE: Creates all the PCST files formatted properly.
'   USE:    Open validation form which will be used to make PCST files. Run program and follow on screen prompts.
' AUTHOR: Jonathan Adams
'
'
'

Dim wb As Workbook
Dim FirstWkbk As Workbook
Dim sht As Worksheet
Dim Sheet As Worksheet
Dim tbl As ListObject
Dim Validation_File_Name As Variant
Dim cValue As Variant
Dim Val_Wk_Array As Variant
Dim Val_Tbl_Name_Array As Variant
Dim CS_72_Header_Array As Variant
Dim Others_Header_Array As Variant
Dim Clin_Doc_Col_Num_Array As Variant
Dim Clin_Doc_Col_Ltr_Array As Variant
Dim Unmapped_Col_Ltr_Array As Variant
Dim Unmapped_Col_Num_Array As Variant
Dim CS_72_Header_Temp_Array As Variant
Dim Others_Header_Temp_Array As Variant
Dim StartCell As Range
Dim rList As Range
Dim User_Name As String
Dim Project_Name As String
Dim Save_Path As String
Dim Code_Sheet As String
Dim Current_Value As String
Dim Sheet_Name As String
Dim CurrentSheet As String
Dim Source_Combined As String
Dim Current_Source As String
Dim Name_Input_Checker As Integer
Dim Confirm_Scrubbed As Integer
Dim Folder_Check As Integer
Dim Project_Name_Checker As Integer
Dim Sources_Check As Integer
Dim LastRow As Long
Dim LastColumn As Long
Dim Next_Blank_Row As Long
Dim Table_ObjIsVisible As Boolean
Dim Checker_Health_Maint As Boolean



    ' 'This disables settings to improve macro performance.
    ' Application.ScreenUpdating = False
    ' Application.Calculation = xlCalculationManual
    ' Application.EnableEvents = False


    ' DEBUG
    ' User_Name = "ja052464"
    ' Project_Name = "Test"
    ' Validation_File_Name = ActiveWorkbook.Name

    ' Error Handling
    ' On Error GoTo ErrHandler

    Val_Wk_Array = Array("Clinical Documentation", "Unmapped Codes", "Health Maintenance Summary")
    Val_Tbl_Name_Array = Array("Clinical_Table", "Unmapped_Table", "Health_Maint_Table")

    CS_72_Header_Array = Array("Registry", "Measure", "Concept", "Source", "DocumentType", "Name", "Section", "DTA", "EventCode", "EventDisplay", "ESH", "ControlType", "NomenclatureID", "Nomenclature", "TaskAssay", "Notes", "Comments", "Standard Code", "Standard Coding System")
    CS_72_Header_Temp_Array = Array("Registry", "Measure", "Concept", "Source", "DocumentType", "Name", "Section", "DTA", "EventCode", "EventDisplay", "ESH", "ControlType", "NomenclatureID", "Nomenclature", "TaskAssay", "Notes", "Comments", "Standard Code", "Standard Coding System")

    Others_Header_Array = Array("Registry", "Measure", "Concept", "Source", "DocumentType", "Name", "Section", "DTA", "Code", "Display", "ESH", "ControlType", "NomenclatureID", "Nomenclature", "vlookup", "Team", "Comments", "Standard Code", "Standard Coding System")
    Others_Header_Temp_Array = Array("Registry", "Measure", "Concept", "Source", "DocumentType", "Name", "Section", "DTA", "Code", "Display", "ESH", "ControlType", "NomenclatureID", "Nomenclature", "vlookup", "Team", "Comments", "Standard Code", "Standard Coding System")

    Clin_Doc_Col_Num_Array = Array("Source", "EventCode", "EventDisplay", "NomenclatureID", "Nomenclature")
    Clin_Doc_Col_Ltr_Array = Array("Source", "EventCode", "EventDisplay", "NomenclatureID", "Nomenclature")

    Unmapped_Col_Ltr_Array = Array("Source", "Code System", "Raw Code", "Raw Display", "Code Short Name")
    Unmapped_Col_Num_Array = Array("Source", "Code System", "Raw Code", "Raw Display", "Code Short Name")

    'This disables settings to improve macro performance.
    ' Application.ScreenUpdating = False
    ' Application.Calculation = xlCalculationManual
    ' Application.EnableEvents = False

    'Prompts user to confirm they have reviewed the data in the validation form BEFORE running this.
    Confirm_Scrubbed = MsgBox("*NOTICE* It is highly advised that you review the data on the Unmapped Codes, Clinical Documentation, and the Health Maintenance Summary Sheet before running this program." & vbNewLine & vbNewLine & "You should delete unneeded lines and review concept endings to confirm the data is correct before proceeding. Otherwise errors will multiplied accross all newly created files.", vbOKCancel + vbQuestion, "Empty Sheet")

    'If user hits cancel then close program.
    If Confirm_Scrubbed = vbCancel Then
        MsgBox ("Program is canceling per user action.")
        GoTo End_Program

    End If

    'Tells user the program is about to to and to leave their machine alone until the program is completed.
    MsgBox ("Program is about to run. This will take a minute or two to complete. During the process you will be asked for a couple inputs at the beginning and you may receive popups asking for permission to overwrite files if you have already ran this program once before. Otherwise please leave computer alone until completed popup appears.")

    'Names variable current file name
    Validation_File_Name = ActiveWorkbook.Name

    'Checks to confirm the user entered a correct user ID. This is needed for file save path.
    Do
        Name_Input_Checker = 0
        User_Name = InputBox("Please enter your Cerner userID." & vbNewLine & vbNewLine & "ex. BE042983")

        If User_Name = vbNullString Then
            MsgBox ("Canceling program per user action.")
            GoTo End_Program

        ElseIf Len(User_Name) <> 8 Then
            MsgBox ("That was not the correct format...Lets try this again..." & vbNewLine & "Please enter your user_ID. No spaces" & vbNewLine & vbNewLine & "ex. BE042983")

        Else
            Name_Input_Checker = 1
        End If

    Loop While Name_Input_Checker = 0


    'Checks to confirm user entered correct project name. This is needed for file name.
    Do
        Project_Name_Checker = 0
        Project_Name = InputBox("Please enter the abbreviation for this project." & vbNewLine & vbNewLine & "ex. NBRO")

        If Project_Name = vbNullString Then
            MsgBox ("Canceling program per user action.")
            GoTo End_Program

        ElseIf Len(Project_Name) = 4 Or Len(Project_Name) = 7 Then    'If length of user inut incorrect, prompt user to try again.
            Project_Name_Checker = 1
        Else
            MsgBox ("Lets try this again.... Please enter the project name..." & vbNewLine & vbNewLine & "ex. NBRO")
        End If

    Loop While Project_Name_Checker = 0

    'Assigns file save path to variable.
    Save_Path = "C:\Users\" & User_Name & "\Documents\" & Project_Name & "_" & "PCST_Files"


    'If the folder already exists then do nothing. Else make it.
    If Len(Dir(Save_Path, vbDirectory)) = 0 Then
        MkDir Save_Path    'Creates the folder

    Else
        Folder_Check = MsgBox("Looks like the folder already exists... Do you want to continue?", vbOKCancel + vbQuestion, "Empty Sheet")    'Folder already exists so continuing on.
    End If

    'If user hits cancel on the folder check then cancel program.
    If Folder_Check = vbCancel Then
        MsgBox ("Canceling program per user action.")
    End If


    ' PRIMARY - Formats worksheets for copying to new workbook
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    For i = 0 To UBound(Val_Wk_Array)

        Sheets(Val_Wk_Array(i)).Select

        ' table can not be created if autofilters are on
        If ActiveSheet.AutoFilterMode = True Then
            ActiveSheet.AutoFilterMode = False
        End If

        ' Checks the current sheet. If it is in table format, convert it to range.
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

        End If


        'Converts range to table.
        Set tbl = ActiveSheet.ListObjects.Add(xlSrcRange, Selection, , xlYes)
          tbl.Name = Val_Tbl_Name_Array(i)
          tbl.TableStyle = "TableStyleLight12"

        'Filters to remove blank lines
        ActiveSheet.ListObjects(1).Range.AutoFilter Field:=5, _
                Criteria1:="<>"

    Next i


    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'PRMARY - CREATES THE SOURCE CODE SHEET AND TABLE FOR LOOP ON THE VALIDATION FORM
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
    For i = 0 To UBound(Val_Wk_Array)
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
        End If
        'Finds next blank row to add additional sources.
        Next_Blank_Row = Sheets("Sources List").Range("A" & Rows.Count).End(xlUp).Row + 1
    Next i


    ' Create named range of the sources
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


    '
    ' SUB - Shows list of all the sources and has user confirm the sources are correct
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    For Each Source_Name In Range("Sources_List")
        Current_Source = Source_Name
        Source_Combined = Source_Combined & Current_Source & vbNewLine

    Next Source_Name

    Sources_Check = MsgBox("Program found the following sources. Please confirm that all sources are unique and there are no duplicates. If there are please click Cancel, rename the sources and then re-run the program. If the sources are good to go click OK to continue." & vbNewLine & vbNewLine & "Sources List:" & vbNewLine & Source_Combined, vbOKCancel + vbQuestion, "Empty Sheet")

    'If user hits cancel then close program.
    If Sources_Check = vbCancel Then
        MsgBox ("Program is canceling per user action.")
        GoTo End_Program
    End If


    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' PRIMARY - Performs Checks on the Validation Form to Confirm Format is Correct Before Proceeding Any Further
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    ' Checks the Health Maintenance Summary If format is incorrect, then tell user and end program
    If LCase(Sheets("Health Maintenance Summary").Range("K5").Value) <> "source" Then

        MsgBox ("Program has detected a possible error with the Validation Form layout" & vbNewLine & vbNewLine & _
                "Program expected Column K on the Health Maintenance Summary sheet to be 'Source'. Please resolve the issue and then run again.")
        GoTo End_Program
    End If


    '
    ''''''''''''''''''''''''''''''''''''''''
    '   PRIMARY - Create New Workbook
    ''''''''''''''''''''''''''''''''''''''''
    '


    ' Loop through the sources
    For Each Source_Name In Range("Sources_List")
        Set wb = Workbooks.Add    'Opens a new workbook

        ' Saves the new workbook
        With NewBook
            ChDir "C:\Users\" & User_Name & "\Documents\" & Project_Name & "_" & "PCST_Files"
            ActiveWorkbook.SaveAs Filename:= _
                    "C:\Users\" & User_Name & "\Documents\" & Project_Name & "_" & "PCST_Files\" & Source_Name
        End With

        ' Selects new workbook
        Windows(Source_Name & ".xlsx").Activate

        ' Populates basic sheets on new workbook
        With ActiveWorkbook
            .Sheets.Add(After:=.Sheets(.Sheets.Count)).Name = "Clinical Documentation"
            .Sheets.Add(After:=.Sheets(.Sheets.Count)).Name = "Unmapped Codes"
            .Sheets.Add(After:=.Sheets(.Sheets.Count)).Name = "Health Maintenance Summary"
            .Sheets.Add(After:=.Sheets(.Sheets.Count)).Name = "Source_Code_Systems"
        End With

        ' SUB - Copies the data from the sheets in the validation form to the new workbook sheets
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        ' Finds the start cell to begin range for copying. Different rules for Health Maint Sheet
        For i = 0 To UBound(Val_Wk_Array)
            CurrentSheet = Val_Wk_Array(i)
            CurrentTable = Val_Tbl_Name_Array(i)

            Windows(Validation_File_Name).Activate
            Sheets(CurrentSheet).Select

            ' Finds the location of the Registry Column
            If CurrentSheet <> Val_Wk_Array(2) Then
                Range("A2:K2").Name = "Header_row"

                For Each cell In Range("Header_row")
                    If cell = "Registry" Then
                        start_cell = cell.Address
                        Exit For
                    End If
                Next cell

            Else
                ' Changes range for the Health Maintenance Summary Sheet
                Range("A5:K5").Name = "Header_row"

                For Each cell In Range("Header_row")
                    If cell = "EXPECT_NAME" Then
                        start_cell = cell.Address
                        Exit For
                    End If
                Next cell

            End If

            ' Copies the data to the new sheet
            Sheets(CurrentSheet).Select
            ActiveSheet.ListObjects(1).Range.AutoFilter Field:=5, _
                    Criteria1:="<>"
            Range(start_cell).Select
            Range(Selection, Selection.End(xlDown)).Select
            Range(Selection, Selection.End(xlToRight)).Select
            Selection.Copy

            ' Selects the newly created excel file and pastes copied cells onto unmapped codes sheet
            Windows(Source_Name & ".xlsx").Activate
            Sheets(CurrentSheet).Select
            Range("A1").Select
            Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                    :=False, Transpose:=False

            ' Sets header for code short name column
            If CurrentSheet = Val_Wk_Array(1) Then
              Range("K1").Value = "Code Short Name"
            End If
        Next i

        ' Formats the new sheets as tables
        For i = 0 To UBound(Val_Wk_Array)
            ' Confirms Selection of new workbook
            Windows(Source_Name & ".xlsx").Activate
            Sheets(Val_Wk_Array(i)).Select

            ' Sets all cells on sheet to a table.
            Set sht = Worksheets(Val_Wk_Array(i))
            Set StartCell = Range("A1")

            ' Find Last Row and Column
            LastRow = StartCell.SpecialCells(xlCellTypeLastCell).Row
            LastColumn = StartCell.SpecialCells(xlCellTypeLastCell).Column

            'Select Range
            sht.Range(StartCell, sht.Cells(LastRow, LastColumn)).Select

            ' Converts range to table.
            Set tbl = ActiveSheet.ListObjects.Add(xlSrcRange, Selection, , xlYes)
            tbl.Name = Val_Tbl_Name_Array(i)
            tbl.TableStyle = "TableStyleLight12"

            ' Filters to remove blank lines
            ActiveSheet.ListObjects(1).Range.AutoFilter Field:=5, _
                    Criteria1:="<>"

        Next i

        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        ' PRIMARY - Finds Location of headers for the sheets
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''


        ' Finds column Header Locations for Unmapped Columns
        Sheets(Val_Wk_Array(1)).Select
        Range("A1:Q1").Name = "Header_row"

        For i = 0 to UBound(Unmapped_Col_Ltr_Array)
        col_count = 0
          For each header in Range("Header_row")
            col_count = col_count + 1
            If LCase(Unmapped_Col_Ltr_Array(i)) = LCase(header) Then
              Unmapped_Col_Ltr_Array(i) = Mid(header.Address, 2, 1)
              Unmapped_Col_Num_Array(i) = col_count
              exit For
            End If
          Next header
        Next i


        ' Finds Column Header Locations for Clinical Documentation
        Sheets(Val_Wk_Array(0)).Select
        Range("A1:Q1").Name = "Header_row"

        For i = 0 to UBound(Clin_Doc_Col_Num_Array)
        col_Count = 0
          For each header in Range("Header_row")
          col_Count = col_Count + 1
            If LCase(Clin_Doc_Col_Num_Array(i)) = LCase(header) Then
              Clin_Doc_Col_Ltr_Array(i) = Mid(header.Address, 2, 1)
              Clin_Doc_Col_Num_Array(i) = col_Count
              exit For
            End If
          Next header
        Next i


        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        ' PRIMARY - Unmapped Remove Duplicates and Set Code Short Name
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''

        ' Unampped codes removes duplicates by Source, EvCode, and Code Display
        Sheets(Val_Wk_Array(1)).Range(Val_Tbl_Name_Array(1) & "[#All]").RemoveDuplicates Columns:=Array(Unmapped_Col_Num_Array(0), Unmapped_Col_Num_Array(2), Unmapped_Col_Num_Array(3)), Header:=xlYes

        ''
        ' SUB - Code Short Name Creation
        ''''''''''''''''''''''''''''''''
        Set Table_Obj = Sheets(Val_Wk_Array(1)).ListObjects(1)

        'Checks current table to determine if any cells are visible. If cells are visible then set "Table_ObjIsVisible" = TRUE
        Set tbl = Sheets(Val_Wk_Array(1)).ListObjects(1)

        If tbl.Range.SpecialCells(xlCellTypeVisible).Areas.Count > 1 Then
            Table_ObjIsVisible = True
        Else:
            Table_ObjIsVisible = tbl.Range.SpecialCells(xlCellTypeVisible).Rows.Count > 1
        End If

        'If there are unmapped codes, copy the code id to the code id short name column
        If Table_ObjIsVisible = True Then
            Sheets(Val_Wk_Array(1)).select
            Range(Unmapped_Col_Ltr_Array(1) & "2:" & Unmapped_Col_Ltr_Array(1) & Cells.SpecialCells(xlCellTypeLastCell).Row).Copy Range(Unmapped_Col_Ltr_Array(4) & "2")
        End If

        ' Clears any autofilters
        Sheets(Val_Wk_Array(1)).AutoFilter.ShowAllData


        ' filters by concept -> registry for easy reviewing in final product
        With ActiveWorkbook.Sheets(Val_Wk_Array(1)).ListObjects(1).Sort
            .SortFields.Add Key:=Range(Val_Tbl_Name_Array(1) & "[Concept]"), SortOn:= _
                    xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
            .SortFields.Add Key:=Range(Val_Tbl_Name_Array(1) & "[Measure]"), SortOn:= _
                    xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
            .SortFields.Add Key:=Range(Val_Tbl_Name_Array(1) & "[Registry]"), SortOn:= _
                    xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
            .Apply
        End With

        '''''''''''''''''''''''''''''''''''''''''''''
        'PRIMARY - Change Format of Code Short Name
        '''''''''''''''''''''''''''''''''''''''''''''

        'Selects Code Short Column and applies name to range
        Sheets(Val_Wk_Array(1)).Range(Unmapped_Col_Ltr_Array(4)&"2:" & Unmapped_Col_Ltr_Array(4) & ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row).Name = "Code_Short"

        'Assigns range to variable
        Set Rng = Range("Code_Short")

        For Each cell In Rng
            If InStr(cell, "urn:cerner:coding:codingsystem:nomenclature.source_vocab:") > 0 Then  'If cell contains "x"
                cValue = cell.Value
                cPlace = InStr(cell, "vocab")
                cell.Value = "nomenclature - " & Right(cValue, Len(cValue) - (cPlace + 5))    'Replace cell with nomenclature and all text after vocab

                'Checks if code id contains PTCARE and shortens appropriately
            ElseIf InStr(cell, "PTCARE") > 0 Then
                cValue = cell.Value
                cPlace = InStr(cValue, "vocab")
                cell.Value = Right(cValue, Len(cValue) - (cPlace + 5))

                'Checks if code id contains healthmaintenance and then shortens appropriately
            ElseIf InStr(cell, "healthmaintenance") > 0 Then
                cValue = cell.Value
                cPlace = InStr(cValue, "healthmaintenance")
                cell.Value = Right(cValue, Len(cValue) - (cPlace + 16))

                'Checks if code id is normal cerner code set and then shortens appropriately
            ElseIf InStr(cell, "urn:cerner:coding:codingsystem:codeset:") > 0 Then
                cValue = cell.Value
                cPlace = InStr(cValue, "codeset:")
                cell.Value = Right(cValue, Len(cValue) - (cPlace + 7))

                'Checks Catches alternate nomenclature code. This catches the general nomenclature code id's which do not contain the tail descriptor
            ElseIf InStr(cell, "urn:cerner:coding:codingsystem:nomenclature") > 0 Then
                cValue = cell.Value
                cPlace = InStr(cValue, "system:")
                cell.Value = Right(cValue, Len(cValue) - (cPlace + 6))
            End If
        Next cell


        '   Converts Nomenclature - PTCARE to Nomenclature - Patient Care to standardize naming convention.
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

        For Each cell In Rng
            If cell = "nomenclature - PTCARE" Then
              cell.Value = "Nomenclature - Patient Care"
            End If
        Next cell


        ' Checks Code Short length and if it is more than >28 characters then shorten the name.
        For Each Code_Short In Range("Code_Short")
            Current_Value = Code_Short.Value
            If (Len(Current_Value) > 28) Then
                Code_Short.ClearContents
                Code_Short.Value = Left(Current_Value, 28)
            End If

            ' Checks code short name for invalid characters and replaces them with a space
            If InStr(Current_Value, ":") > 0 Then
                Code_Short.ClearContents
                New_Value = Replace(Current_Value, ":", " ")
                Code_Short.Value = New_Value
            End If
        Next Code_Short



        ' SUB - Clinical Documentation Remove Duplicates
        '''''''''''''''''''''''''''''''''''''''''''''

        ' Removes dups by source, evcode, evcode display,
        Sheets(Val_Wk_Array(0)).Range("Clinical_Table[#All]").RemoveDuplicates Columns:=Array(Clin_Doc_Col_Num_Array(0), _
                Clin_Doc_Col_Num_Array(1), Clin_Doc_Col_Num_Array(2)), Header:=xlYes

        ' Removes dups by source, nomenclature ID, Nomenclature Display
        Sheets(Val_Wk_Array(0)).Range("Clinical_Table[#All]").RemoveDuplicates Columns:=Array(Clin_Doc_Col_Num_Array(0), _
                Clin_Doc_Col_Num_Array(3), Clin_Doc_Col_Num_Array(4)), Header:=xlYes


        ' SUB - Creates the source code worksheet
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Windows(Source_Name & ".xlsx").Activate

        Sheets(Val_Wk_Array(1)).ListObjects("Unmapped_Table").Range.AutoFilter Field:=4, _
                Criteria1:=Source_Name, Operator:=xlAnd

        ' Pastes the values from the Code Short Name onto the Source_Code_Systems Sheet
        Sheets(Val_Wk_Array(1)).Range(Unmapped_Col_Ltr_Array(4) & "1:" & Unmapped_Col_Ltr_Array(4) & ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row).Copy Sheets("Source_Code_Systems").Range("A1")

        Sheets("Source_Code_Systems").Select

        ' Formats Source_Code_Systems Sheet
        Set sht = Worksheets("Source_Code_Systems")
        Set StartCell = Range("A1")

        ' Refresh UsedRange
        Worksheets("Source_Code_Systems").UsedRange

        ' Find Last Row and Column
        LastRow = StartCell.SpecialCells(xlCellTypeLastCell).Row
        LastColumn = StartCell.SpecialCells(xlCellTypeLastCell).Column

        ' Select Range
        sht.Range(StartCell, sht.Cells(LastRow, LastColumn)).Select

        Set tbl = ActiveSheet.ListObjects.Add(xlSrcRange, Selection, , xlYes)
          tbl.Name = "Code_ID_Table"
          tbl.TableStyle = "TableStyleLight9"

          ' Removes Duplicates from the Code ID Lists
        Range("Code_ID_Table[[#Headers],[Code Short Name]]").Select
        Application.CutCopyMode = False
        ActiveSheet.Range("Code_ID_Table[#All]").RemoveDuplicates Columns:=1, Header:= _
                xlYes

        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        ' PRIMARY - Checks to determine how many unique code ID's there are for this source.
        ' If there is only 1 source "72" then set range to one cell, otherwise select all cells and set the range.
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


        'If there are no unmapped codes for this source, then set code source to 72 only.
        If Range("A2") = "" Then
            Range("A2").Value = "72"
        End If

        ' If A3 is empty and thus only 1 code id, then set the range of Code_ID_List to just cell A2
        If Range("A3") = "" Then
            Range("A2").Name = "Code_ID_List"
        Else
        Range("A2").Select
        Range(Selection, Selection.End(xlDown)).Select
        Selection.Name = "Code_ID_List"
        End If

        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        '      PRIMARY - Creates a new sheet and names the sheet with the current source
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

        'Loops through each unique Code ID for this source and creates a sheet with the relavant data.
        For Each code In Range("Code_ID_List")

            With ActiveWorkbook
                .Sheets.Add(After:=.Sheets(.Sheets.Count)).Name = code
            End With

            Code_Sheet = code

            'Special instructions for code set 72
            If code = "72" Then


                Off_Count = 0
                For i = 0 To UBound(CS_72_Header_Array)
                    ' Uses cell "A1" as the starting point for the header row.
                    Sheets(Code_Sheet).Range("A1").Offset(0, Off_Count).Value = CS_72_Header_Array(i)    ' places next array value within the next column
                    Off_Count = Off_Count + 1    'Increases the offset count on each loop
                Next i

                ' SUB - Copies Clinical Documentation to 72
                '''''''''''''''''''''''''''''''''''''

                'Filters for only the current source
                Sheets(Val_Wk_Array(0)).ListObjects("Clinical_Table").Range.AutoFilter Field:=Clin_Doc_Col_Num_Array(0), _
                        Criteria1:=Source_Name, Operator:=xlAnd


                'Checks current table to determine if any cells are visible. If cells are visible then set "Table_ObjIsVisible" = TRUE
                Set tbl = Sheets(Val_Wk_Array(1)).ListObjects(1)

                If tbl.Range.SpecialCells(xlCellTypeVisible).Areas.Count > 1 Then
                    Table_ObjIsVisible = True
                Else:
                    Table_ObjIsVisible = tbl.Range.SpecialCells(xlCellTypeVisible).Rows.Count > 1
                End If

                If tbl.Range.SpecialCells(xlCellTypeVisible).Areas.Count > 1 Then
                    Table_ObjIsVisible = True
                Else
                    Table_ObjIsVisible = False

                End If

                'If data is visible, then copy visible data
                If Table_ObjIsVisible = True Then

                    ' Copies Registry - TaskAssay columns
                    Sheets(Val_Wk_Array(0)).Select
                    Range("A1:O1").Select
                    Range(Selection, Selection.End(xlDown)).Copy Sheets(Code_Sheet).Range("A1")

                    ' Copies the Notes Column
                    Sheets(Val_Wk_Array(0)).Select
                    Range("P1").Select
                    Range(Selection, Selection.End(xlDown)).Copy Sheets(Code_Sheet).Range("P1")

                    ' Copies the Team Column
                    Sheets(Val_Wk_Array(0)).Select
                    Range("Q1").Select
                    Range(Selection, Selection.End(xlDown)).Copy Sheets(Code_Sheet).Range("O1")

                End If


                '    SUB - Copies unmapped codes to 72 sheet
                '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                Sheets(Val_Wk_Array(1)).Select


                ' Applies filters for only this source and code being currently reviewed.
                ActiveSheet.ListObjects("Unmapped_Table").Range.AutoFilter Field:=Unmapped_Col_Num_Array(0), _
                        Criteria1:=Source_Name, Operator:=xlAnd
                ActiveSheet.ListObjects("Unmapped_Table").Range.AutoFilter Field:=Unmapped_Col_Num_Array(4), _
                        Criteria1:=code, Operator:=xlAnd

                Set tbl = ActiveSheet.ListObjects(1)

                If tbl.Range.SpecialCells(xlCellTypeVisible).Areas.Count > 1 Then
                    Table_ObjIsVisible = True
                Else:
                    Table_ObjIsVisible = tbl.Range.SpecialCells(xlCellTypeVisible).Rows.Count > 1
                End If

                If tbl.Range.SpecialCells(xlCellTypeVisible).Areas.Count > 1 Then
                    Table_ObjIsVisible = True
                Else
                    Table_ObjIsVisible = False

                End If
                ' If data is visible then copy data
                If Table_ObjIsVisible = True Then

                    ' Finds the next blank row on the code sheet
                    Sheets(Code_Sheet).Select
                    Next_Blank_Row = Range("A" & Rows.Count).End(xlUp).Row + 1

                    ' Copies Registry through Code System ID to new sheet
                    Sheets(Val_Wk_Array(1)).Select
                    Range("A2:E2").Select
                    Range(Selection, Selection.End(xlDown)).Copy Sheets(Code_Sheet).Range("A" & Next_Blank_Row)

                    ' Copies Raw Code through Count
                    Sheets(Val_Wk_Array(1)).Select
                    Range("F2:H2").Select
                    Range(Selection, Selection.End(xlDown)).Copy Sheets(Code_Sheet).Range("I" & Next_Blank_Row)

                    ' Copies unmapped Code Notes Column
                    Sheets(Val_Wk_Array(1)).Select
                    Range("I2").Select
                    Range(Selection, Selection.End(xlDown)).Copy Sheets(Code_Sheet).Range("P" & Next_Blank_Row)

                    ' Copies unmapped Team Column
                    Sheets(Val_Wk_Array(1)).Select
                    Range("J2").Select
                    Range(Selection, Selection.End(xlDown)).Copy Sheets(Code_Sheet).Range("O" & Next_Blank_Row)

                End If


                '       SUB - Populates Health Maintenance visible data to CS 72
                '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

                Sheets("Health Maintenance Summary").Select

                ActiveSheet.ListObjects("Health_Maint_Table").Range.AutoFilter Field:=11, _
                        Criteria1:=Source_Name, Operator:=xlAnd

                Set tbl = ActiveSheet.ListObjects(1)

                ' Checks table for visible data
                If tbl.Range.SpecialCells(xlCellTypeVisible).Areas.Count > 1 Then
                    Table_ObjIsVisible = True
                Else:
                    Table_ObjIsVisible = tbl.Range.SpecialCells(xlCellTypeVisible).Rows.Count > 1
                End If

                ' If data is visible then copy visible data
                If Table_ObjIsVisible = True Then

                    'Copies Sources Column
                    Range("K2:K" & Cells.SpecialCells(xlCellTypeLastCell).Row).Select
                    Selection.Copy
                    Sheets(Code_Sheet).Select

                    'Uses The CodeSet 72 column "Sources" to determine next blank row
                    Next_Blank_Row = Range("D" & Rows.Count).End(xlUp).Row + 1

                    ' Pastes Sources on new sheet
                    Range("D" & Next_Blank_Row).Select
                    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                            :=False, Transpose:=False

                    ' Copies Expect_Meaning Column
                    Sheets("Health Maintenance Summary").Select
                    Range("B2:B" & Cells.SpecialCells(xlCellTypeLastCell).Row).Select
                    Selection.Copy
                    Sheets(Code_Sheet).Select

                    'Pastes Expect_Meaning on new sheet
                    Range("F" & Next_Blank_Row).Select
                    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                            :=False, Transpose:=False

                    ' Copies Satisfier_Meaning Column
                    Sheets("Health Maintenance Summary").Select
                    Range("G2:G" & Cells.SpecialCells(xlCellTypeLastCell).Row).Select
                    Selection.Copy
                    Sheets(Code_Sheet).Select

                    ' Pastes Satisfier_Meaning on new sheet
                    Range("G" & Next_Blank_Row).Select
                    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                            :=False, Transpose:=False

                    ' Copies Entry_Type Column
                    Sheets("Health Maintenance Summary").Select
                    Range("C2:C" & Cells.SpecialCells(xlCellTypeLastCell).Row).Select
                    Selection.Copy
                    Sheets(Code_Sheet).Select

                    ' Pastes Entry_Type on new sheet
                    Range("L" & Next_Blank_Row).Select
                    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                            :=False, Transpose:=False

                    ' Copies Event_CD Column
                    Sheets("Health Maintenance Summary").Select
                    Range("I2:I" & Cells.SpecialCells(xlCellTypeLastCell).Row).Select
                    Selection.Copy
                    Sheets(Code_Sheet).Select

                    ' Pastes Event_CD on new sheet
                    Range("I" & Next_Blank_Row).Select
                    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                            :=False, Transpose:=False

                    ' Copies Event_CD_DISP Column
                    Sheets("Health Maintenance Summary").Select
                    Range("J2:J" & Cells.SpecialCells(xlCellTypeLastCell).Row).Select
                    Selection.Copy
                    Sheets(Code_Sheet).Select

                    ' Pastes Event_CD_DISP on new sheet
                    Range("J" & Next_Blank_Row).Select
                    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                            :=False, Transpose:=False
                End If


                '       SUB - Populates headers for all other sheets
                '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

            Else

                Off_Count = 0
                For i = 0 To UBound(Others_Header_Array)
                    ' Uses cell "A1" as the starting point for the header row.
                    Range("A1").Offset(0, Off_Count).Value = Others_Header_Array(i)    ' places next array value within the next column
                    Off_Count = Off_Count + 1    'Increases the offset count on each loop
                Next i


                ' Filters unmapped codes table for current source and code within loop.
                Sheets(Val_Wk_Array(1)).Select
                ActiveSheet.ListObjects("Unmapped_Table").Range.AutoFilter Field:=Unmapped_Col_Num_Array(0), _
                        Criteria1:=Source_Name, Operator:=xlAnd
                ActiveSheet.ListObjects("Unmapped_Table").Range.AutoFilter Field:=Unmapped_Col_Num_Array(4), _
                        Criteria1:=code, Operator:=xlAnd

                ' Sets variable to the table on the active sheet.
                Set tbl = ActiveSheet.ListObjects(1)

                If tbl.Range.SpecialCells(xlCellTypeVisible).Areas.Count > 1 Then
                    Table_ObjIsVisible = True
                Else:
                    Table_ObjIsVisible = tbl.Range.SpecialCells(xlCellTypeVisible).Rows.Count > 1
                End If

                If tbl.Range.SpecialCells(xlCellTypeVisible).Areas.Count > 1 Then
                    Table_ObjIsVisible = True
                Else
                    Table_ObjIsVisible = False

                End If
                ' If data is visible then copy data.
                If Table_ObjIsVisible = True Then

                    ' Finds next blank row
                    Sheets(Code_Sheet).Select
                    Next_Blank_Row = Range("A" & Rows.Count).End(xlUp).Row + 1

                    ' Copies Registry through Code System ID to new sheet
                    Sheets(Val_Wk_Array(1)).Select
                    Range("A2:E2").Select
                    Range(Selection, Selection.End(xlDown)).Copy Sheets(Code_Sheet).Range("A" & Next_Blank_Row)

                    ' Copies Raw Code through Count
                    Sheets(Val_Wk_Array(1)).Select
                    Range("F2:H2").Select
                    Range(Selection, Selection.End(xlDown)).Copy Sheets(Code_Sheet).Range("I" & Next_Blank_Row)

                    ' Copies unmapped Code Notes Column
                    Sheets(Val_Wk_Array(1)).Select
                    Range("I2").Select
                    Range(Selection, Selection.End(xlDown)).Copy Sheets(Code_Sheet).Range("P" & Next_Blank_Row)

                    ' Copies unmapped Team Column
                    Sheets(Val_Wk_Array(1)).Select
                    Range("J2").Select
                    Range(Selection, Selection.End(xlDown)).Copy Sheets(Code_Sheet).Range("O" & Next_Blank_Row)
                End If

            End If

        Next code

        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        '   PRIMARY - Populates Nomenclature - Patient Care Sheet If Needed
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


        Sheets(Val_Wk_Array(0)).Select

        'Filters table for current source.
        Sheets(Val_Wk_Array(0)).ListObjects("Clinical_Table").Range.AutoFilter Field:=Clin_Doc_Col_Num_Array(0), _
                Criteria1:=Source_Name, Operator:=xlAnd

        ' Filters to remove lines without nomenclature
        Sheets(Val_Wk_Array(0)).ListObjects("Clinical_Table").Range.AutoFilter Field:=13, _
                Criteria1:="<>"

        ' Eventually update this to filter out all rows which ARE MAPPED CORRECTLY. To only leave
        ' incorrect rows.
        ' ActiveSheet.ListObjects("Clinical_Table").Range.AutoFilter Field:=18, _
          '  Criteria1:="<>"

        Set tbl = Sheets(Val_Wk_Array(0)).ListObjects(1)

        ' Checks filtered table for visible data.
        If tbl.Range.SpecialCells(xlCellTypeVisible).Areas.Count > 1 Then
            Table_ObjIsVisible = True
        Else
            Table_ObjIsVisible = False
        End If

        ' If data is visible then copy data.
        If Table_ObjIsVisible = True Then
            ' Check to see if Nomenclature - Patient Care sheet already exists
            For Each Sheet In Worksheets
                If Sheet.Name = "Nomenclature - Patient Care" Then
                    exists = True
                    Exit For
                Else
                    exists = False
                End If
            Next Sheet

            ' If sheet does NOT exist, then create the sheet
            If exists = False Then
                ActiveWorkbook.Sheets.Add(After:=Worksheets(1)).Name = "Nomenclature - Patient Care"

                Off_Count = 0
                For i = 0 To UBound(Others_Header_Array)
                    ' Uses cell "A1" as the starting point for the header row.
                    Range("A1").Offset(0, Off_Count).Value = Others_Header_Array(i)    ' places next array value within the next column
                    Off_Count = Off_Count + 1
                Next i
            End If

            ' Populates the Nomenclature - Patient Care Sheet with data from Clinical Documentation
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            ' Finds next Available Row
            Sheets("Nomenclature - Patient Care").Select
            Next_Blank_Row = Range("A" & Rows.Count).End(xlUp).Row + 1

            Sheets(Val_Wk_Array(0)).Select
            Range("A2:O2").Select
            Range(Selection, Selection.End(xlDown)).Copy Sheets("Nomenclature - Patient Care").Range("A" & Next_Blank_Row)

            ' Copies the Notes / Comments Column
            Sheets(Val_Wk_Array(0)).Select
            Range("P2").Select
            Range(Selection, Selection.End(xlDown)).Copy Sheets("Nomenclature - Patient Care").Range("Q" & Next_Blank_Row)

            ' Copies the Team Column
            Sheets(Val_Wk_Array(0)).Select
            Range("Q2").Select
            Range(Selection, Selection.End(xlDown)).Copy Sheets("Nomenclature - Patient Care").Range("P" & Next_Blank_Row)


            Set StartCell = Range("A1")
            Set sht = Worksheets("Nomenclature - Patient Care")

            ' Finds last row with text
            LastRow = StartCell.SpecialCells(xlCellTypeLastCell).Row


            ' Removes duplicates from sheet by source, nom ID and Nom description
            ' Sheets("Nomenclature - Patient Care").Range("$A$1:$O$" & LastRow).RemoveDuplicates Columns:=Array(4, 13, 14), _
                    ' Header:=xlYes

        End If

        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        '   PRIMARY - Workbook Cleanup
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


        Application.DisplayAlerts = False
        ' Deletes the extra sheets not needed
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

        Application.DisplayAlerts = True

        '
        ' SUB - Creates and Populates Index Sheet
        '''''''''''''''''''''''''''''''''''''''''

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

        '
        ' SUB - Formats remaining sheets as tables for appearance
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''

        For Each Sheet In Worksheets
            Sheet.Activate

            Set sht = Sheet    ' Sets value
            Set StartCell = Range("A1")    ' Start cell used to determine where to begin creating the table range

            LastRow = StartCell.SpecialCells(xlCellTypeLastCell).Row
            LastColumn = StartCell.SpecialCells(xlCellTypeLastCell).Column
            Sheet_Name = Sheet.Name    ' Assigns sheet name to a variable as a string

            sht.Range(StartCell, sht.Cells(LastRow, LastColumn)).Select

            Set tbl = ActiveSheet.ListObjects.Add(xlSrcRange, Selection, , xlYes)
            tbl.Name = Sheet_Name
            tbl.TableStyle = "TableStyleLight9"
            Columns.AutoFit
            Range("A1").Select

        Next Sheet

        '
        '  SUB - Aligns Index Sheet
        '''''''''''''''''''''''

        Sheets("INDEX SHEET").Select
        Range("A2").Select
        Range(Selection, Selection.End(xlDown)).Select
        With Selection
            .HorizontalAlignment = xlLeft
        End With
        Range("A2").Select

        '
        '  SUB - Saves the new workbook
        ''''''''''''''''''''''''''''''''''''''

        Workbooks(Source_Name & ".xlsm").Close SaveChanges:=True
        Windows(Validation_File_Name).Activate    'Switches back to old workbook to begin next loop

    Next Source_Name    'Start over with next source from list



End_Program:
  'Re-enables previously disabled settings after all code has run.
  Application.ScreenUpdating = True
  Application.Calculation = xlCalculationAutomatic
  Application.EnableEvents = True

  ' Clears the filesystem descriptor allowing you to delete the folder
  Dir "C:\"

  'Notifies user that the program has completed.
  MsgBox ("Your PCST Files have been created. Folder is loctated within your My Documents.")
Exit Sub

ErrHandler:

  'Re-enables previously disabled settings after all code has run.
  Application.ScreenUpdating = True
  Application.Calculation = xlCalculationAutomatic
  Application.EnableEvents = True

  ' Clears the filesystem descriptor allowing you to delete the folder
  Dir "C:\"
  MsgBox("Program encountered an error " & vbCrLf & Err.Description & vbCrLf & DisplayErr & vbCrLf & Err.Number)

End Sub
