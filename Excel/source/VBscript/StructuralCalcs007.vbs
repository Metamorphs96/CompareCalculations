Option Explicit

Sub MainApplicationV7()
  Dim struct1(6)
  Dim i
  
  For i = 0 To UBound(struct1)
    Set struct1(i) = New PhysicalCharacteristic
    struct1(i).initialise
  Next
  
  'Input Parameters
  Call struct1(0).setCharacteristic("Cpe", -0.7, "", 1)
  Call struct1(1).setCharacteristic("qz", 0.96, "kPa", 2)
  Call struct1(2).setCharacteristic("s", 3, "m", 2)
  Call struct1(3).setCharacteristic("L", 6, "m", 2)
    
  'Results
  Call struct1(4).setCharacteristicDefinition("pn", "kPa", 2)
  Call struct1(5).setCharacteristicDefinition("w", "kN/m", 2)
  Call struct1(6).setCharacteristicDefinition("M", "kNm",2)
  
  struct1(4).Value = struct1(0).Value * struct1(1).Value
  struct1(5).Value = struct1(4).Value * struct1(2).Value
  struct1(6).Value = struct1(5).Value * struct1(3).Value ^ 2 / 8

  'Report
  For i = 0 To UBound(struct1)
    With struct1(i)
      .cprint
    End With
  Next

End Sub
