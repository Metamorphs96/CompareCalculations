Option Explicit

Function setStructuralCharacteristic(xName, xUnits, xFormatStr)
  Dim tmpStruct
  
  Set tmpStruct = New PhysicalCharacteristic
  With tmpStruct
    .initialise
    Call .setCharacteristicDefinition(xName, xUnits, xFormatStr)
  End With
  
  Set setStructuralCharacteristic = tmpStruct
  
End Function

Sub MainApplicationV8()
  Dim Cpe
  Dim qz
  Dim s
  Dim L
  
  Dim pn
  Dim w
  Dim M
  
  'Input Parameters
  Set Cpe = setStructuralCharacteristic("Cpe", "", 1)
  Set qz = setStructuralCharacteristic("qz", "kPa", 2)
  Set s = setStructuralCharacteristic("s", "m", 2)
  Set L = setStructuralCharacteristic("L", "m", 2)
  
  Cpe.Value = -0.7
  qz.Value = 0.96
  s.Value = 3
  L.Value = 6
    
  'Results
  Set pn = setStructuralCharacteristic("pn", "kPa", 2)
  Set w = setStructuralCharacteristic("w", "kN/m", 2)
  Set M = setStructuralCharacteristic("M", "kNm", 2)
  
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

