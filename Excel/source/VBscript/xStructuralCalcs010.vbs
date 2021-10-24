Option Explicit

class PhysicalCharacteristic
	Public key 
	Public Name 
	Public Value
	Public Units
	Public formatStr

	Sub initialise()
	  key = 0
	  Name = ""
	  Value = 0
	  Units = ""
	  formatStr = "0.00"
	End Sub

	Sub setCharacteristicKeyed(xkey,xName , xValue, xUnits , xFormatStr)
	  
	  key = xkey
	  Name = xName
	  Value = xValue
	  Units = xUnits
	  formatStr = xFormatStr
	  
	End Sub

	Sub setCharacteristic(xName , xValue, xUnits, xFormatStr )
	  
	  Name = xName
	  Value = xValue
	  Units = xUnits
	  formatStr = xFormatStr
	  
	End Sub

	Sub setCharacteristicDefinition(xName , xUnits, xFormatStr)
	  
	  Name = xName
	  Units = xUnits
	  formatStr = xFormatStr
	  
	End Sub


	Function sprint()
		sprint = Name & ": " & FormatNumber(Value, formatStr) & " " & Units
	End Function

	Sub cprint()
		WScript.Echo sprint
	End Sub
End Class

Function calcNetPressure(Cpe, qz ) 
  calcNetPressure = Cpe * qz
End Function

Function calcUDL(pn, s )
  calcUDL = pn * s
End Function

Function calcMoment(w, L )
  calcMoment = w * L ^ 2 / 8
End Function

Function setStructuralCharacteristic(xName, xUnits, xFormatStr) 
  Dim tmpStruct
  
  Set tmpStruct = New PhysicalCharacteristic
  With tmpStruct
    .initialise
    Call .setCharacteristicDefinition(xName, xUnits, xFormatStr)
  End With
  
  Set setStructuralCharacteristic = tmpStruct
  
End Function



Sub MainApplicationV10()
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
  
  pn.Value = calcNetPressure(Cpe.Value, qz.Value)
  w.Value = calcUDL(pn.Value, s.Value)
  M.Value = calcMoment(w.Value, L.Value)

  'Report
  For Each key In struct01.Keys
    Set tmpStruct = struct01.Item(key)
    tmpStruct.cprint
  Next
  
  WScript.Echo "--------------"
  struct01.Item("M").cprint
  struct01("M").cprint
  WScript.Echo struct01("Cpe").Value
  
End Sub
