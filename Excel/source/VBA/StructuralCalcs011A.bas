Attribute VB_Name = "StructuralCalcs011A"
Option Explicit

Sub MainApplicationV11()
  Dim Cpe As Double
  Dim qz As Double
  Dim s As Double
  Dim L As Double
  
  Dim pn As Double
  Dim w As Double
  Dim M As Double
  
  Dim key As Variant
  Dim tmpStruct As PhysicalCharacteristic
  
  Dim struct01 As Dictionary
  
  Set struct01 = New Dictionary
  
  'Input Parameters
  Set tmpStruct = setStructuralCharacteristic("Cpe", "", "0.0")
  Call struct01.Add(tmpStruct.Name, tmpStruct)
  Set tmpStruct = setStructuralCharacteristic("qz", "kPa", "0.00")
  Call struct01.Add(tmpStruct.Name, tmpStruct)
  Set tmpStruct = setStructuralCharacteristic("s", "m", "0.00")
  Call struct01.Add(tmpStruct.Name, tmpStruct)
  Set tmpStruct = setStructuralCharacteristic("L", "m", "0.00")
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
  Set tmpStruct = setStructuralCharacteristic("pn", "kPa", "0.00")
  Call struct01.Add(tmpStruct.Name, tmpStruct)
  Set tmpStruct = setStructuralCharacteristic("w", "kN/m", "0.00")
  Call struct01.Add(tmpStruct.Name, tmpStruct)
  Set tmpStruct = setStructuralCharacteristic("M", "kNm", "0.00")
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
  Next key
  
    
End Sub
