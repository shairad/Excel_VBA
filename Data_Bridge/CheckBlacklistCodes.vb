Sub CheckBlacklistCodes()

Dim ws As Worksheet
Dim Header_Check As Boolean
Dim UnmappedCodeIdHeader As Variant
Dim Header_User_Response As Variant
Dim SheetArray As Variant
Dim HeaderLocations As Variant
Dim HeaderNames As Variant

SheetArray = Array("Unmapped Codes", "BlackList_Table")
HeaderLocations = Array("Unmapped Location", "Blacklist Location")
HeaderNames = Array("Coding System ID", "BlacklistedCodeSystem")



' SUB - Finds the column for the code system ID on the unmapped codes Sheet
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

For i = 0 to UBound(SheetArray)
' Sets header row range
Sheets(SheetArray(i)).Select
Range("A1").Select
Range("A1", Selection.End(xlToRight)).Name = "Header_row"

' Finds the Code System ID header column

Header_Check = False
  For each header in Range("Header_row")
    If LCase(header) = LCase("Coding System ID") Or LCase(header) = LCase("BlacklistedCodeSystem") Then
      HeaderLocations(i) = Mid(header.Address, 2, 1)
      Header_Check = True
      exit For
    End If
  next header

  ' If no header was found then prompt the user for the column or allow the user to cancel the program
  If Header_Check = False Then
      Header_User_Response = InputBox("BORIS was unable to find the header:" & vbNewLine & HeaderNames(i) & " on the " & SheetArray(i) & vbNewLine & vbNewLine & "BORIS needs your help. Please enter the letter of the column containing the Coding System ID's.","If I am BORIS who are you?")
      If Header_User_Response = vbNullString Then
          GoTo User_Exit
      Else
          HeaderLocations(i) = UCase(UHeader_User_Response)
      End If
  End If
Next i


' SUB - Finds blacklist column on the imported table








End Sub
