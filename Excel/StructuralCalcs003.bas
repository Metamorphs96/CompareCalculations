Attribute VB_Name = "StructuralCalcs003"
Option Explicit

Function getBendingMoment(Cpe As Double, qz As Double, s As Double, L As Double) As Double
  Dim pn As Double
  Dim w As Double
  
  pn = Cpe * qz     'kPa
  w = pn * s
  getBendingMoment = w * L ^ 2 / 8 'kNm
  
End Function

Sub MainApplicationV3()
  Dim Cpe As Double
  Dim qz As Double
  Dim s As Double
  Dim L As Double
  Dim M As Double
  
  Cpe = -0.7
  qz = 0.96         'kPa
  s = 3             'm
  L = 6             'm
  
  M = getBendingMoment(Cpe, qz, s, L)
  Debug.Print "Moment: " & Format(M, "0.00") & " kNm"
End Sub

