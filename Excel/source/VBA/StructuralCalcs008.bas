Attribute VB_Name = "StructuralCalcs008"
Option Explicit

Function setStructuralCharacteristic(xName As String, xUnits As String, xFormatStr As String) As PhysicalCharacteristic
  Dim tmpStruct As PhysicalCharacteristic
  
  Set tmpStruct = New PhysicalCharacteristic
  With tmpStruct
    .initialise
    Call .setCharacteristicDefinition(xName, xUnits, xFormatStr)
  End With
  
  Set setStructuralCharacteristic = tmpStruct
  
End Function

Sub MainApplicationV8()
  Dim Cpe As PhysicalCharacteristic
  Dim qz As PhysicalCharacteristic
  Dim s As PhysicalCharacteristic
  Dim L As PhysicalCharacteristic
  
  Dim pn As PhysicalCharacteristic
  Dim w As PhysicalCharacteristic
  Dim M As PhysicalCharacteristic
  
  'Input Parameters
  Set Cpe = setStructuralCharacteristic("Cpe", "", "0.0")
  Set qz = setStructuralCharacteristic("qz", "kPa", "0.00")
  Set s = setStructuralCharacteristic("s", "m", "0.00")
  Set L = setStructuralCharacteristic("L", "m", "0.00")
  
  Cpe.Value = -0.7
  qz.Value = 0.96
  s.Value = 3
  L.Value = 6
    
  'Results
  Set pn = setStructuralCharacteristic("pn", "kPa", "0.00")
  Set w = setStructuralCharacteristic("w", "kN/m", "0.00")
  Set M = setStructuralCharacteristic("M", "kNm", "0.00")
  
  pn.Value = Cpe.Value * qz.Value
  w.Value = pn.Value * s.Value
  M.Value = w.Value * L.Value ^ 2 / 8

  'Report
  Cpe.cprint
  qz.cprint
  s.cprint
  L.cprint
  pn.cprint
  w.cprint
  M.cprint

End Sub

