Sub Testing()
Dim DataRange As Variant
Dim Irow As Long
Dim Icol As Integer
Dim MyType As Variant
Dim MyNote As Variant


DataRange = Range("Data_Range").Value

For Irow = 1 To UBound(DataRange)
  For Icol = 1
    MyType = DataRange(Irow, Icol)

    If MyType =  "PowerForm" Then
      DataRange(Irow, 13) = "Yes"
    End If
  Next Icol
Next Irow
Range("Data_Range").Value = DataRange

End Sub
