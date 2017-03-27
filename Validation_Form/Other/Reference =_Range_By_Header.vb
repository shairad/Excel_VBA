'The below example searches for the header based on the name. In the example, it colors the entire 'column Red based on match, but you can change that. The variables should be easy to follow, but 'please let me know of questions.

Sub Color_Range_Based_On_Header()
    Dim rngHeaders As Range
    Dim rngHdrFound As Range

    Const ROW_HEADERS As Integer = 1
    Const HEADER_NAME As String = "Location"

    Set rngHeaders = Intersect(Worksheets("Sheet1").UsedRange, Worksheets("Sheet1").Rows(ROW_HEADERS))
    Set rngHdrFound = rngHeaders.Find(HEADER_NAME)

    If rngHdrFound Is Nothing Then
        'Do whatever you want if the header is missing
        Exit Sub
    End If

    Range(rngHdrFound, rngHdrFound.End(xlDown)).Interior.Color = vbRed

End Sub
