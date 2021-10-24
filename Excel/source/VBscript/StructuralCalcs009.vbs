Option Explicit


Sub MainApplicationV9()
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
  Set Cpe = setStructuralCharacteristic("Cpe", "", 1)
  Set qz = setStructuralCharacteristic("qz", "kPa", 2)
  Set s = setStructuralCharacteristic("s", "m", 2)
  Set L = setStructuralCharacteristic("L", "m", 2)
  
  Call struct01.Add(Cpe.Name, Cpe)
  Call struct01.Add(qz.Name, qz)
  Call struct01.Add(s.Name, s)
  Call struct01.Add(L.Name, L)
  
  Cpe.Value = -0.7
  qz.Value = 0.96
  s.Value = 3
  L.Value = 6
    
  'Results
  Set pn = setStructuralCharacteristic("pn", "kPa", 2)
  Set w = setStructuralCharacteristic("w", "kN/m", 2)
  Set M = setStructuralCharacteristic("M", "kNm", 2)
  
  Call struct01.Add(pn.Name, pn)
  Call struct01.Add(w.Name, w)
  Call struct01.Add(M.Name, M)
  
  pn.Value = Cpe.Value * qz.Value
  w.Value = pn.Value * s.Value
  M.Value = w.Value * L.Value ^ 2 / 8

  'Report
  For Each key In struct01.Keys
    Set tmpStruct = struct01.Item(key)
    tmpStruct.cprint
  Next
  
  WScript.Echo "--------------"
  struct01.Item("M").cprint
  struct01("M").cprint

End Sub


