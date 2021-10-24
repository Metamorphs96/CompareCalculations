Option Explicit

Sub MainApplicationV11B()
  Dim Cpe
  Dim qz
  Dim s
  Dim L
  
  Dim pn
  Dim w
  Dim M
  
  Dim key
  Dim tmpStruct
  
  Dim struct01
  
  Set struct01 = CreateObject("Scripting.Dictionary")
  
  'Set Up Data Store
  'Input Parameters
  Set tmpStruct = setStructuralCharacteristic("Cpe", "", 1)
  Call struct01.Add(tmpStruct.Name, tmpStruct)
  Set tmpStruct = setStructuralCharacteristic("qz", "kPa", 2)
  Call struct01.Add(tmpStruct.Name, tmpStruct)
  Set tmpStruct = setStructuralCharacteristic("s", "m", 2)
  Call struct01.Add(tmpStruct.Name, tmpStruct)
  Set tmpStruct = setStructuralCharacteristic("L", "m", 2)
  Call struct01.Add(tmpStruct.Name, tmpStruct)
  'Results
  Set tmpStruct = setStructuralCharacteristic("pn", "kPa", 2)
  Call struct01.Add(tmpStruct.Name, tmpStruct)
  Set tmpStruct = setStructuralCharacteristic("w", "kN/m", 2)
  Call struct01.Add(tmpStruct.Name, tmpStruct)
  Set tmpStruct = setStructuralCharacteristic("M", "kNm", 2)
  Call struct01.Add(tmpStruct.Name, tmpStruct)
  
  'Do Calculations
  Cpe = -0.7
  qz = 0.96
  s = 3
  L = 6
  pn = Cpe * qz
  w = pn * s
  M = w * L ^ 2 / 8
  
  
  
  
  'Store Data Values
  struct01("Cpe").Value = Cpe
  struct01("qz").Value = qz
  struct01("s").Value = s
  struct01("L").Value = L
  struct01("pn").Value = pn
  struct01("w").Value = w
  struct01("M").Value = M
  

  'Report
  For Each key In struct01.Keys
    Set tmpStruct = struct01.Item(key)
    tmpStruct.cprint
  Next
  
    
End Sub


