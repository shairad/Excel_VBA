Sub CheckPrevCodeSystems()

Dim ws As Worksheet
Dim Header_Check As Boolean
Dim EvCodeHeader As Variant
Dim Header_User_Response As Variant
Dim SheetArray As Variant
Dim HeaderLocations As Variant
Dim HeaderNames As Variant
Dim EvCodeCheck As Variant
Dim EvCodeCheckAnswerArray As Variant
Dim UnmappedEvCodeArray As Variant
Dim EvCodeCheckHeader As String
Dim EvCodeConcat As String
Dim LR As Long
Dim Lookup As Variant
Dim cell_Lookup As Variant
Dim sResult_Value As Variant

SheetArray = Array("Unmapped Codes", "CodeSystemCheck")
HeaderLocations = Array("Unmapped Location", "Raw Code", "Blacklist Location")
HeaderNames = Array("Coding System ID", "Raw Code", "EventCodeCheck")
UnmappedAddHeaders = Array("EvCodeCheck", "CodeLookup")

' TODO - Write the code whic concats the code system ID and the RawCode column
' TODO - Then Pass that column to memory and handle vlookup and writing out results


' SUB - Finds the column for the code system ID on the unmapped codes Sheet
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

For i = 0 to UBound(SheetArray)
' Sets header row range
Sheets(SheetArray(i)).Select
Range("A1").Select
Range("A2", Selection.End(xlToRight)).Name = "Header_row"

' Finds the Code System ID header column

Header_Check = False
  For each header in Range("Header_row")
    If LCase(header) = LCase("Coding System ID") _
     Or LCase(header) = LCase("Raw Code") _
      Or LCase(header) = LCase("EventCodeCheck") Then
      HeaderLocations(i) = Mid(header.Address, 2, 1)
      Header_Check = True
      exit For
    End If
  next header

  ' If no header was found then prompt the user for the column or allow the user to cancel the program
  If Header_Check = False Then
      Header_User_Response = InputBox("BORIS was unable to find the header:" & vbNewLine & HeaderNames(i) & " on the " & SheetArray(i) & vbNewLine & vbNewLine & "BORIS needs your help. Please enter the letter of the missing column.","If I am BORIS who are you?")
      If Header_User_Response = vbNullString Then
          ' GoTo User_Exit
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

For i = 0 to UBound(UnmappedAddHeaders)
  Header_Check = False
  For each header in Range("Header_row")
    If LCase(header) = LCase(UnmappedAddHeaders(i)) then
      UnmappedAddHeaders(i) = Mid(header.Address, 2, 1)
      Header_Check = True
      exit For
    End If
  next header
  ' If the column does not exist, then create it.
  If Header_Check = False Then
    NextBlank = Mid(Cells(2, Columns.Count).End(xlToLeft).Offset(0,1).Address,2,1)
    Range(NextBlank & "2") = UnmappedAddHeaders(i)
    UnmappedAddHeaders(i) = NextBlank
  End If
Next i




  ' SUB - Assigns Unmapped Codes Code System ID's to array in memory
  ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  Sheets(SheetArray(0)).Select
  LR = Range(HeaderLocations(0) & Rows.Count).End(xlUp).Row
  Range(HeaderLocations(0)&"3:"  & HeaderLocations(0) & LR).SpecialCells(xlCellTypeVisible).Name = "UnmappedCodeID_Range"

  UnmappedEvCodeArray = Range("UnmappedCodeID_Range").Value

  ' SUB - Set EvCodeCheck answer range to array in memory
  Sheets(SheetArray(0)).Select
  Range(EvCodeCheckHeader &"3:" & EvCodeCheckHeader & LR).SpecialCells(xlCellTypeVisible).Name = "EvCodeChecker"

  EvCodeCheckAnswerArray = Range("EvCodeChecker")


  ' SUB - Assigns Blacklist Codes from the BlackList Table to an array in memory
  '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  Sheets(SheetArray(1)).Select
  LR = Range(HeaderLocations(1) & Rows.Count).End(xlUp).Row
  Range(HeaderLocations(1) &"3:" & HeaderLocations(1) & LR).SpecialCells(xlCellTypeVisible).Name = "PreviousEvCodes"

  BlacklistArray = Range("PreviousEvCodes").Value



  For i = 1 to UBound(UnmappedEvCodeArray)

      cell_Lookup = UnmappedEvCodeArray(i,1)  'The cell value you want to look for a match with

      sResult = Application.VLookup(cell_Lookup, Range("BlackList_Range"), 1, False)

      If IsError(sResult) Then
          sResult_Value = "0"
          EvCodeCheckAnswerArray(i,1) = sResult_Value
      Else
          EvCodeCheckAnswerArray(i,1) = "On Blacklist"

      End If

  Next i

  'Write the updated DataRange Array to the excel file
  Range("BlackListAnswers").Value = EvCodeCheckAnswerArray

  Sheets(SheetArray(0)).Select
  MsgBox ("BORIS has completed the Blacklist Code System check")


End Sub
