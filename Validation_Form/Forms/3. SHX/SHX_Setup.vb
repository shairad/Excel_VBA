Sub SHX_Setup()
'
' SHX REsults Macro. Formats range as table. Inserts lookup formulas and populates autofill.
'
Dim tbl As ListObject
Dim sht As Worksheet
Dim LastRow As Long
Dim LastColumn As Long
Dim StartCell As Range
Dim rList As Range
Dim Header_Array As Variant

    'Disables settings to improve performance
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False

    Header_Array = Array("NOMEN_ID", "CS_72", "CS_14003", "CS_4002165")

    MsgBox ("Program is about to run. Please leave computer alone until completed")

    Sheets("Social History Results").Select

    ActiveSheet.AutoFilterMode = False        'Disables autoFilter

    'If table exists on sheet then convert to range
    If ActiveSheet.ListObjects.Count > 0 Then

        With ActiveSheet.ListObjects(1)
            Set rList = .Range
            .Unlist
        End With

        With rList
            .Interior.ColorIndex = xlColorIndexNone
            .Font.ColorIndex = xlColorIndexAutomatic
            .Borders.LineStyle = xlLineStyleNone
        End With

    End If

    Set sht = Worksheets("Social History Results")
    Set StartCell = Range("A1")

    'Refresh UsedRange
    Worksheets("Social History Results").UsedRange

    'Find Last Row and Column
    LastRow = StartCell.SpecialCells(xlCellTypeLastCell).Row
    LastColumn = StartCell.SpecialCells(xlCellTypeLastCell).Column

    'Select Range
    sht.Range(StartCell, sht.Cells(LastRow, LastColumn)).Select

    'Turn selected Range Into Table
    Set tbl = ActiveSheet.ListObjects.Add(xlSrcRange, Selection, , xlYes)
    tbl.Name = "SHX_Results"
    tbl.TableStyle = "TableStyleLight12"

    'changes font color of header row to white
    Rows("1:1").Select
    With Selection.Font
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = 0
    End With

    Sheets("Social History Results").Select

    Range("A1").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Name = "Header_Row"


    ' Finds the locations of the lookup columns
    For Each Header In Range("Header_Row")
        For i = 0 To UBound(Header_Array)
            If Header_Array(i) = Header Then
                Header_Array(i) = Mid(Header.Address, 2, 1)
                Exit For
            End If
        Next i
    Next Header


    ' Checks Validated Mappings Code ID Column to confirm format is NumberFormat
    Sheets("Validated Mappings").Select

    Range("A1").Select
    Range(Selection, Selection.End(xlToRight)).Name = "Val_Headers"

    For Each cell In Range("Val_Headers")
        If cell = LCase("Code ID") Then
            CodeID = Mid(cell.Address, 2, 1)
        End If
        Exit For

        ' Sets Code ID range for looping to check format
        Range(CodeID & "2:" & CodeID & ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row).Name = "CodeID_Column"

        For Each cell In Range("CodeID_Column")
            If IsNumeric(cell) Then
                cell.Value = Val(cell.Value)
                cell.NumberFormat = "0"
            End If
        Next cell


        ' Checks code columns and makes sure values are in number format for lookup
        For i = 0 To UBound(Header_Array)
            'Selects all cells not empty in the code column and assigns named range for loop
            Range(Header_Array(i) & "2:" & Header_Array(i) & ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row).Name = "Code_Column"

            For Each cell In Range("Code_Column")
                If IsNumeric(cell) Then
                    cell.Value = Val(cell.Value)
                    cell.NumberFormat = "0"
                End If
            Next cell
        Next i


        ' CS Nomenclature Mapped
        Range("A2").Select
        ActiveCell = "=IFERROR(INDEX('Validated Mappings'!I:I,MATCH(" & Header_Array(0) & "2,'Validated Mappings'!D:D,0))," & Header_Array(0) & "2)"

        ' CS 72 formula
        Range("B2").Select
        ActiveCell = "=IFERROR(INDEX('Validated Mappings'!I:I,MATCH(" & Header_Array(1) & "2,'Validated Mappings'!D:D,0))," & Header_Array(1) & "2)"

        ' CS 14003 Formula
        Range("C2").Select
        ActiveCell = "=IFERROR(INDEX('Validated Mappings'!I:I,MATCH(" & Header_Array(2) & "2,'Validated Mappings'!D:D,0))," & Header_Array(2) & "2)"

        ' CS 4002165
        Range("D2").Select
        ActiveCell = "=IFERROR(INDEX('Validated Mappings'!I:I,MATCH(" & Header_Array(3) & "2,'Validated Mappings'!D:D,0))," & Header_Array(3) & "2)"


        'Centers cell values
        Columns("A:D").Select
        With Selection
            .HorizontalAlignment = xlCenter
            .Orientation = 0
            .AddIndent = False
            .IndentLevel = 0
            .ShrinkToFit = False
            .ReadingOrder = xlContext
            .MergeCells = False
        End With

        'Re-enables Auto-calculate for forumlas
        Application.Calculation = xlCalculationAutomatic

        Sheets("Social History Results").Select
        Cells.Select
        Selection.Copy
        Sheets("To_Review").Select
        Range("A1").Select
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                :=False, Transpose:=False

        ActiveSheet.AutoFilterMode = False

        'If table exists on this sheet, then convert to range
        If ActiveSheet.ListObjects.Count > 0 Then

            With ActiveSheet.ListObjects(1)
                Set rList = .Range
                .Unlist        ' convert the table back to a range
            End With

            With rList
                .Interior.ColorIndex = xlColorIndexNone
                .Font.ColorIndex = xlColorIndexAutomatic
                .Borders.LineStyle = xlLineStyleNone
            End With

        End If


        Set sht = Worksheets("To_Review")
        Set StartCell = Range("A1")

        'Refresh UsedRange
        Worksheets("Social History Results").UsedRange

        'Find Last Row and Column
        LastRow = StartCell.SpecialCells(xlCellTypeLastCell).Row
        LastColumn = StartCell.SpecialCells(xlCellTypeLastCell).Column

        'Select Range
        sht.Range(StartCell, sht.Cells(LastRow, LastColumn)).Select

        Set tbl = ActiveSheet.ListObjects.Add(xlSrcRange, Selection, , xlYes)
        tbl.Name = "SHX_To_Review"
        tbl.TableStyle = "TableStyleLight9"

        Cells.Select
        Cells.EntireColumn.AutoFit
        Cells.Select
        Cells.EntireRow.AutoFit

        Columns("A:D").Select
        With Selection
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .Orientation = 0
            .AddIndent = False
            .IndentLevel = 0
            .ShrinkToFit = False
            .ReadingOrder = xlContext
            .MergeCells = False
        End With

        Range("A1").Select

        Application.ScreenUpdating = True
        Application.EnableEvents = True

        MsgBox ("Program Completed")

    End Sub
