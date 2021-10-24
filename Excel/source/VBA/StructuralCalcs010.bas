Attribute VB_Name = "StructuralCalcs010"
Option Explicit

Function calcNetPressure(Cpe As Double, qz As Double) As Double
  calcNetPressure = Cpe * qz
End Function

Function calcUDL(pn As Double, s As Double) As Double
  calcUDL = pn * s
End Function

Function calcMoment(w As Double, L As Double) As Double
  calcMoment = w * L ^ 2 / 8
End Function

Sub MainApplicationV10()
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
  
  pn.Value = calcNetPressure(Cpe.Value, qz.Value)
  w.Value = calcUDL(pn.Value, s.Value)
  M.Value = calcMoment(w.Value, L.Value)

  'Report
  For Each key In struct01.Keys
    Set tmpStruct = struct01.Item(key)
    tmpStruct.cprint
  Next key
  
  Debug.Print "--------------"
  struct01.Item("M").cprint
  struct01("M").cprint
  Debug.Print struct01("Cpe").Value
  
  Debug.Print Application.Evaluate(Cpe.Value * qz.Value)
  
End Sub
