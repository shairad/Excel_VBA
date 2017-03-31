
Sub Color_Range_Based_On_Header()

  Dim Sheet_Headers As Variant
  Dim Find_Header As Range
  Dim rngHeaders As Range
  Dim ColHeaders As Variant
  Dim EV_Code_Map_Col As Variant
  Dim Nom_Map_Col As Variant
  Dim Both_Map_Col As Variant
  Dim DTA_Col As Variant
  Dim Nom_Source_Col As Variant
  Dim Alpha_Nom_Col As Variant



  Range("A1").Select
  Range(Selection, Selection.End(xlToRight)).Select

  Selection.Name = "Header_Row"

  Header_Names = Array("Event Code Mapped?", "Nomenclature Mapped?", "Both Validated?", "DTA_EC")


  For each cell in Range("Header_Row")

    If cell = "Event Code Mapped?" Then
      EV_Code_Map_Col = Mid(cell.Address, 2, 1)

    Elseif cell = "Nomenclature Mapped?" Then
      Nom_Map_Col = Mid(cell.Address, 2, 1)

    Elseif cell = "Both Validated?" Then
      Both_Map_Col = Mid(cell.Address, 2, 1)

    Elseif cell = "DTA_EC" Then
      DTA_Col = Mid(cell.Address, 2, 1)

    Elseif cell = "NOMEN_SOURCESRG" Then
      Nom_Source_Col = Mid(cell.Address, 2, 1)

    Elseif cell = "ALPHA_NOMEN_ID" Then
      Alpha_Nom_Col = Mid(cell.Address, 2, 1)

    End if

  Next cell


End Sub
