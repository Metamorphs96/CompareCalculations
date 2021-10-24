Attribute VB_Name = "StructuralCalcs009"
Option Explicit


Sub MainApplicationV9()
  Dim Cpe As PhysicalCharacteristic
  Dim qz As PhysicalCharacteristic
  Dim s As PhysicalCharacteristic
  Dim L As PhysicalCharacteristic
  
  Dim pn As PhysicalCharacteristic
  Dim w As PhysicalCharacteristic
  Dim M As PhysicalCharacteristic
  
  Dim key As Variant
  Dim tmpStruct As PhysicalCharacteristic
  
  Dim struct01 As Dictionary
  
  Set struct01 = New Dictionary
  
  'Input Parameters
  Set Cpe = setStructuralCharacteristic("Cpe", "", "0.0")
  Set qz = setStructuralCharacteristic("qz", "kPa", "0.00")
  Set s = setStructuralCharacteristic("s", "m", "0.00")
  Set L = setStructuralCharacteristic("L", "m", "0.00")
  
  Call struct01.Add(Cpe.Name, Cpe)
  Call struct01.Add(qz.Name, qz)
  Call struct01.Add(s.Name, s)
  Call struct01.Add(L.Name, L)
  
  Cpe.Value = -0.7
  qz.Value = 0.96
  s.Value = 3
  L.Value = 6
    
  'Results
  Set pn = setStructuralCharacteristic("pn", "kPa", "0.00")
  Set w = setStructuralCharacteristic("w", "kN/m", "0.00")
  Set M = setStructuralCharacteristic("M", "kNm", "0.00")
  
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
  Next key
  
  Debug.Print "--------------"
  struct01.Item("M").cprint
  struct01("M").cprint

End Sub


