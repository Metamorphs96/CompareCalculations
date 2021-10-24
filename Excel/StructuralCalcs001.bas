Attribute VB_Name = "StructuralCalcs001"
Sub MainApplicationV1()
  Cpe = -0.7
  qz = 0.96   'kPa
  pn = Cpe * qz 'kPa
  w = pn * 3  'kN/m
  L = 6       'm
  M = w * L ^ 2 / 8 'kNm
  Debug.Print "Moment: ", FormatNumber(M, 2)
End Sub
