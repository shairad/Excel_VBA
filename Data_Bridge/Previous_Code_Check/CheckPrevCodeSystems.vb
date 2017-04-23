Sub CheckPrevCodeSystems()

Dim Header_Check As Boolean
Dim Header_User_Response As Variant
Dim SheetArray As Variant
Dim HeaderLocations As Variant
Dim HeaderNames As Variant
Dim EvCodeCheck As String
Dim EvCodeCheckAnswerArray As Variant
Dim EVCodeCheckArray As Variant
Dim EvCodeCheckHeader As String
Dim EvCodeConcat As String
Dim LR As Long
Dim Lookup As Variant
Dim cell_Lookup As Variant
Dim sResult_Value As Variant

    SheetArray = Array("Unmapped Codes", "CodeSystemCheck")
    HeaderLocations = Array("Unmapped Location", "Raw Code", "EventCodeCheck")
    HeaderNames = Array("Coding System ID", "Raw Code", "EventCodeCheck")
    UnmappedHeaders = Array("EvCodeCheck", "CodeLookup")


    ' SUB - Finds the column for the code system ID on the unmapped codes Sheet
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    For i = 0 To UBound(SheetArray)
        ' Sets header row range
        Sheets(SheetArray(i)).Select
        Range("A1").Select
        Range("A2", Selection.End(xlToRight)).Name = "Header_row"

        ' Finds columns by header name
        Header_Check = False
        For Each Header In Range("Header_row")
            If LCase(Header) = LCase("Coding System ID") _
                    Or LCase(Header) = LCase("Raw Code") _
                    Or LCase(Header) = LCase("EventCodeCheck") Then
                HeaderLocations(i) = Mid(Header.Address, 2, 1)
                Header_Check = True
                Exit For
            End If
        Next Header

        ' If no header was found then prompt the user for the column or allow the user to cancel the program
        If Header_Check = False Then
            Header_User_Response = InputBox("BORIS was unable to find the header:" & vbNewLine & HeaderNames(i) & " on the " & SheetArray(i) & vbNewLine & vbNewLine & "BORIS needs your help. Please enter the letter of the missing column.", "If I am BORIS who are you?")
            If Header_User_Response = vbNullString Then
                GoTo User_Exit
            Else
                HeaderLocations(i) = UCase(Header_User_Response)
            End If
        End If
    Next i


    ' SUB - Checks if there already is a column titled EVCodeCheck, if not then create a new one.
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Sheets(SheetArray(0)).Select
    Range("B2").Select
    Range(Selection, Selection.End(xlToRight)).Name = "Header_row"

    For i = 0 To UBound(UnmappedHeaders)
        Header_Check = False
        For Each Header In Range("Header_row")
            If LCase(Header) = LCase(UnmappedHeaders(i)) Then
                UnmappedHeaders(i) = Mid(Header.Address, 2, 1)
                Header_Check = True
                Exit For
            End If
        Next Header
        ' If the column does not exist, then create it.
        If Header_Check = False Then
            NextBlank = Mid(Cells(2, Columns.Count).End(xlToLeft).Offset(0, 1).Address, 2, 1)
            Range(NextBlank & "2") = UnmappedHeaders(i)
            UnmappedHeaders(i) = NextBlank
        End If
    Next i


    ' SUB - Creates Concat Column
    Sheets(SheetArray(0)).Select
    LR = Range(HeaderLocations(0) & Rows.Count).End(xlUp).Row
    Range(UnmappedHeaders(1) & "3:" & UnmappedHeaders(1) & LR).Formula = "=CONCATENATE(F3,""|"",G3)"


    ' SUB - Assigns CodeLookup Column to an array in memory
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Sheets(SheetArray(0)).Select
    LR = Range(HeaderLocations(0) & Rows.Count).End(xlUp).Row
    Range(UnmappedHeaders(1) & "3:" & UnmappedHeaders(1) & LR).SpecialCells(xlCellTypeVisible).Name = "CodeLookup"

    EVCodeCheckArray = Range("CodeLookup").Value

    ' SUB - Set EvCodeCheck answer range to array in memory
    Sheets(SheetArray(0)).Select
    Range(UnmappedHeaders(0) & "3:" & UnmappedHeaders(0) & LR).SpecialCells(xlCellTypeVisible).Name = "EvCodeCheck"

    EvCodeCheckAnswerArray = Range("EvCodeCheck")


    ' SUB - Assigns a name to the Previous ev code lookup column
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Sheets(SheetArray(1)).Select
    LR = Range(HeaderLocations(1) & Rows.Count).End(xlUp).Row
    Range(HeaderLocations(1) & "3:" & HeaderLocations(1) & LR).SpecialCells(xlCellTypeVisible).Name = "PreviousEvCodes"


    ' SUB - Chehcks each cell in the EvCodeCheck for matches and either assigns a match or returns 0
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    For i = 1 To UBound(EVCodeCheckArray)
        cell_Lookup = EVCodeCheckArray(i, 1)
        sResult = Application.VLookup(cell_Lookup, Range("PreviousEvCodes"), 1, False)
        If IsError(sResult) Then
            sResult_Value = "0"
            EvCodeCheckAnswerArray(i, 1) = sResult_Value
        Else
            EvCodeCheckAnswerArray(i, 1) = "Previously Reviewed"
        End If
    Next i

    ' Write the updated DataRange Array to the excel file
    Range("EvCodeCheck").Value = EvCodeCheckAnswerArray

    ' Tells user program is completed
    Sheets(SheetArray(0)).Select
    MsgBox ("BORIS has completed the Blacklist Code System check")
    Exit Sub

User_Exit:
    MsgBox ("Exiting per user action")

End Sub
