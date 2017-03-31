'The below example searches for the header based on the name. In the example, it colors the entire 'column Red based on match, but you can change that. The variables should be easy to follow, but 'please let me know of questions.

Sub Color_Range_Based_On_Header()

  Dim Sheet_Headers As Variant
  Dim rngDateHeader As range
  Dim rngHeaders As range


  Set rngHeaders = range("1:1")
  Set rngDateHeader = rngHeaders.Find("Event Code Mapped?")

  Columnlocation = Mid(rngDateHeader.Address, 2, 1)


End Sub
