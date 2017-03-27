Sub Delete_Extra_Sheets()

Dim sheet As Worksheet

Application.DisplayAlerts = False

  For Each sheet In Worksheets

    If sheet.name = "Unmapped Codes" _
     Or sheet.name = "Health Maintenance Summary" _
     Or sheet.name = "Clinical Documentation" _
     Or sheet.name = "Source_Code_Systems" _
     Or sheet.name = "Sheet1" _
     Then
     sheet.Delete
    End If
  Next sheet

Application.DisplayAlerts = True

End Sub
