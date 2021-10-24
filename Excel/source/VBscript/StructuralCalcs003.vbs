Option Explicit

Function getBendingMoment(Cpe, qz, s, L) 
  Dim pn
  Dim w
  
  pn = Cpe * qz     'kPa
  w = pn * s
  getBendingMoment = w * L ^ 2 / 8 'kNm
  
End Function

Sub MainApplicationV3()
  Dim Cpe
  Dim qz
  Dim s
  Dim L
  Dim M
  
  Cpe = -0.7
  qz = 0.96         'kPa
  s = 3             'm
  L = 6             'm
  
  M = getBendingMoment(Cpe, qz, s, L)
  WScript.Echo "Moment: " & FormatNumber(M, 2) & " kNm"
End Sub

