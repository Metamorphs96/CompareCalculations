Option Explicit

Sub MainApplicationV2()
  Dim Cpe
  Dim qz
  Dim pn
  Dim w
  Dim L
  Dim M
  
  Cpe = -0.7
  qz = 0.96         'kPa
  pn = Cpe * qz     'kPa
  w = pn * 3        'kN/m
  L = 6             'm
  M = w * L ^ 2 / 8 'kNm
  WScript.Echo "Moment: ", FormatNumber(M, 2)
End Sub
