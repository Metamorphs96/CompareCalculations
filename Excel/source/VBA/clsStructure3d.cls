VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsStructure3d"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Cpe As Double
Public qz As Double
Public s As Double
Public L As Double

Dim report(7) As String

Sub initialise()
  Cpe = 0
  qz = 0
  s = 0
  L = 0
End Sub

Sub SetParameters(Cpe1 As Double, qz1 As Double, s1 As Double, L1 As Double)
  Cpe = Cpe1
  qz = qz1
  s = s1
  L = L1
End Sub

Function calcNetPressure() As Double
  calcNetPressure = Cpe * qz
End Function

Function calcUDL() As Double
  calcUDL = calcNetPressure * s
End Function

Function calcMoment() As Double
  calcMoment = calcUDL * L ^ 2 / 8
End Function

Sub sprint()
  report(0) = "Cpe = " & CStr(Cpe)
  report(1) = "qz = " & CStr(qz) & " kPa"
  report(2) = "s = " & CStr(s) & " m"
  report(3) = "L = " & CStr(L) & " m"
  report(4) = "pn = " & FormatNumber(calcNetPressure, 2) & " kPa"
  report(5) = "w = " & FormatNumber(calcUDL, 2) & " kN/m"
  report(6) = "M = " & FormatNumber(calcMoment, 2) & " kNm"
  report(7) = "------------------------------------"
End Sub

Sub cprint()
  Dim s As Variant
  
  sprint
  
  For Each s In report
    Debug.Print s
  Next s
  
End Sub

Sub fprint(fp As Integer)
  Dim s As Variant
  
  sprint
  
  For Each s In report
    Print #fp, s
  Next s
  
End Sub

Sub cprint1()
  Debug.Print "Cpe = " & CStr(Cpe)
  Debug.Print "qz = " & CStr(qz) & " kPa"
  Debug.Print "s = " & CStr(s) & " m"
  Debug.Print "L = " & CStr(L) & " m"
  Debug.Print "pn = " & FormatNumber(calcNetPressure, 2) & " kPa"
  Debug.Print "w = " & FormatNumber(calcUDL, 2) & " kN/m"
  Debug.Print "M = " & FormatNumber(calcMoment, 2) & " kNm"
  Debug.Print "------------------------------------"
End Sub

Sub fprint1(fp As Integer)
  Print #fp, "Cpe = " & CStr(Cpe)
  Print #fp, "qz = " & CStr(qz) & " kPa"
  Print #fp, "s = " & CStr(s) & " m"
  Print #fp, "L = " & CStr(L) & " m"
  Print #fp, "pn = " & FormatNumber(calcNetPressure, 2) & " kPa"
  Print #fp, "w = " & FormatNumber(calcUDL, 2) & " kN/m"
  Print #fp, "M = " & FormatNumber(calcMoment, 2) & " kNm"
  Print #fp, "------------------------------------"
End Sub

