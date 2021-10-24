Option Explicit

Class PhysicalCharacteristic
	Public key 
	Public Name
	Public Value 'String or variant may be more suitable for general case
	Public Units
	Public formatStr

	Sub initialise()
	  key = 0
	  Name = ""
	  Value = 0
	  Units = ""
	  formatStr = 2
	End Sub

	Sub setCharacteristicKeyed(xkey,xName, xValue, xUnits, xFormatStr)
	  
	  key = xkey
	  Name = xName
	  Value = xValue
	  Units = xUnits
	  formatStr = xFormatStr
	  
	End Sub

	Sub setCharacteristic(xName, xValue, xUnits, xFormatStr)
	  
	  Name = xName
	  Value = xValue
	  Units = xUnits
	  formatStr = xFormatStr
	  
	End Sub

	Sub setCharacteristicDefinition(xName, xUnits, xFormatStr)
	  
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