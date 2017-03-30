Sub SpecialLoop()
    Dim cl As Range
    Dim rng As Range


'''Populates notes

  ActiveSheet.ListObjects("New_Lines").range.AutoFilter Field:=5, Criteria1:= _
      "=PowerForm", Operator:=xlOr, Criteria2:="=IView"
  ActiveSheet.ListObjects("New_Lines").range.AutoFilter Field:=20, Criteria1 _
      :="=Validated", Operator:=xlAnd
  ActiveSheet.ListObjects("New_Lines").range.AutoFilter Field:=19, Criteria1 _
      :="<>"Validated"", Operator:=xlAnd

    Range("T1").Select
    range(Selection, Selection.End(xlDown)).Select
    Set rng = Selection

    For Each cl In rng.SpecialCells(xlCellTypeVisible)
        Debug.Print cl
    Next cl

End Sub
