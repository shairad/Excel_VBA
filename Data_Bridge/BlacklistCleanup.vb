Sub BlacklistCleanup()

' Description - Deletes the blacklist code sheet after it has been used to check against for matches

Dim sheet As Worksheets

    Application.DisplayAlerts = False

    For Each sheet In Worksheets
        If sheet.Name = "BlackList_Table" _
                Then
            sheet.Delete
        End If
    Next sheet

    Application.DisplayAlerts = True


End Sub
