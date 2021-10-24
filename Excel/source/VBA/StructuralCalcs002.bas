Attribute VB_Name = "StructuralCalcs002"
Option Explicit

Sub MainApplicationV2()
  Dim Cpe As Double
  Dim qz As Double
  Dim pn As Double
  Dim w As Double
  Dim L As Double
  Dim M As Double
  
  Cpe = -0.7
  qz = 0.96         'kPa
  pn = Cpe * qz     'kPa
  w = pn * 3        'kN/m
  L = 6             'm
  M = w * L ^ 2 / 8 'kNm
  Debug.Print "Moment: ", FormatNumber(M, 2)
End Sub
