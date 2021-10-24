Attribute VB_Name = "StructuralCalcs004"
Option Explicit

Type Structure3D
  Cpe As Double
  qz As Double
  s As Double
  L As Double
  pn As Double
  w As Double
  M As Double
End Type

Function getBendingMomentV2(ByRef struct1 As Structure3D) As Double
  
  With struct1
    .pn = .Cpe * .qz     'kPa
    .w = .pn * .s        'kN/m
    getBendingMomentV2 = .w * .L ^ 2 / 8 'kNm
  End With
  
End Function


Sub MainApplicationV4()
  Dim struct1 As Structure3D
  
  With struct1
    .Cpe = -0.7
    .qz = 0.96
    .s = 3
    .L = 6
    .M = getBendingMomentV2(struct1)
    
    Debug.Print "Cpe = " & CStr(.Cpe)
    Debug.Print "qz = " & CStr(.qz) & " kPa"
    Debug.Print "s = " & CStr(.s) & " m"
    Debug.Print "L = " & CStr(.L) & " m"
    Debug.Print "pn = " & CStr(.pn) & " kPa"
    Debug.Print "w = " & CStr(.w) & " kN/m"
    
    Debug.Print "M = " & FormatNumber(.M, 2) & " kNm"
    Debug.Print "------------------------------------"
  End With
  
End Sub


