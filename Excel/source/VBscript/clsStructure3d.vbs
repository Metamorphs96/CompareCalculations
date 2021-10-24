Option Explicit

Class clsStructure3D
	Public Cpe
	Public qz
	Public s
	Public L

	Dim report(7)

	Sub initialise()
	  Cpe = 0
	  qz = 0
	  s = 0
	  L = 0
	End Sub

	Sub SetParameters(Cpe1, qz1, s1, L1)
	  Cpe = Cpe1
	  qz = qz1
	  s = s1
	  L = L1
	End Sub

	Function calcNetPressure()
	  calcNetPressure = Cpe * qz
	End Function

	Function calcUDL()
	  calcUDL = calcNetPressure * s
	End Function

	Function calcMoment()
	  calcMoment = calcUDL * L ^ 2 / 8
	End Function

	Sub sprint()
	  report(0) = "Cpe = " & CStr(Cpe)
	  report(1) = "qz = " & CStr(qz) & " kPa"
	  report(2) = "s = " & CStr(s) & " m"
	  report(3) = "L = " & CStr(L) & " m"
	  report(4) = "pn = " & FormatNumber(calcNetPressure, 2) & " kPa"
	  report(5) = "w = " & FormatNumber(calcUDL, 2) & " kN/m"
	  report(6) = "M = " & FormatNumber(calcMoment, 2) & " kNm"
	  report(7) = "------------------------------------"
	End Sub

	Sub cprint()
	  Dim s
	  
	  sprint
	  
	  For Each s In report
		WScript.Echo s
	  Next
	  
	End Sub

	Sub fprint(fpText)
	  Dim s
	  
	  sprint
	  
	  For Each s In report
		fpText.WriteLine s
	  Next
	  
	End Sub

	Sub cprint1()
	  WScript.Echo "Cpe = " & CStr(Cpe)
	  WScript.Echo "qz = " & CStr(qz) & " kPa"
	  WScript.Echo "s = " & CStr(s) & " m"
	  WScript.Echo "L = " & CStr(L) & " m"
	  WScript.Echo "pn = " & FormatNumber(calcNetPressure, 2) & " kPa"
	  WScript.Echo "w = " & FormatNumber(calcUDL, 2) & " kN/m"
	  WScript.Echo "M = " & FormatNumber(calcMoment, 2) & " kNm"
	  WScript.Echo "------------------------------------"
	End Sub

	Sub fprint1(fpText)
	  fpText.WriteLine "Cpe = " & CStr(Cpe)
	  fpText.WriteLine "qz = " & CStr(qz) & " kPa"
	  fpText.WriteLine "s = " & CStr(s) & " m"
	  fpText.WriteLine "L = " & CStr(L) & " m"
	  fpText.WriteLine "pn = " & FormatNumber(calcNetPressure, 2) & " kPa"
	  fpText.WriteLine "w = " & FormatNumber(calcUDL, 2) & " kN/m"
	  fpText.WriteLine "M = " & FormatNumber(calcMoment, 2) & " kNm"
	  fpText.WriteLine "------------------------------------"
	End Sub
End Class
