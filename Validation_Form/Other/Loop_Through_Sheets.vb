Sub WorksheetLoop()

   Dim WS_Count As Integer
   Dim I As Integer

   ' Set WS_Count equal to the number of worksheets in the active
   ' workbook.
   WS_Count = ActiveWorkbook.Worksheets.Count

   ' Begin the loop.
   For I = 1 To WS_Count

      ' Insert your code here.
      ' The following line shows how to reference a sheet within
      ' the loop by displaying the worksheet name in a dialog box.
      MsgBox ActiveWorkbook.Worksheets(I).Name

   Next I

End Sub
