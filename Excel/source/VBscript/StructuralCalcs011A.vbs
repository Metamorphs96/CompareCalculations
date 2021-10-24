Option Explicit

Sub MainApplicationV11()
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
  
  'Input Parameters
  Set tmpStruct = setStructuralCharacteristic("Cpe", "", 1)
  Call struct01.Add(tmpStruct.Name, tmpStruct)
  Set tmpStruct = setStructuralCharacteristic("qz", "kPa", 2)
  Call struct01.Add(tmpStruct.Name, tmpStruct)
  Set tmpStruct = setStructuralCharacteristic("s", "m", 2)
  Call struct01.Add(tmpStruct.Name, tmpStruct)
  Set tmpStruct = setStructuralCharacteristic("L", "m", 2)
  Call struct01.Add(tmpStruct.Name, tmpStruct)
      
  Cpe = -0.7
  struct01("Cpe").Value = Cpe
  qz = 0.96
  struct01("qz").Value = qz
  s = 3
  struct01("s").Value = s
  L = 6
  struct01("L").Value = L
  
    
  'Results
  Set tmpStruct = setStructuralCharacteristic("pn", "kPa", 2)
  Call struct01.Add(tmpStruct.Name, tmpStruct)
  Set tmpStruct = setStructuralCharacteristic("w", "kN/m", 2)
  Call struct01.Add(tmpStruct.Name, tmpStruct)
  Set tmpStruct = setStructuralCharacteristic("M", "kNm", 2)
  Call struct01.Add(tmpStruct.Name, tmpStruct)
  
  pn = Cpe * qz
  struct01("pn").Value = pn
  w = pn * s
  struct01("w").Value = w
  M = w * L ^ 2 / 8
  struct01("M").Value = M


  'Report
  For Each key In struct01.Keys
    Set tmpStruct = struct01.Item(key)
    tmpStruct.cprint
  Next
  
    
End Sub

