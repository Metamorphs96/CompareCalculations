'Get Values of Input Parameters
Cpe=InputBox("Pressure Coefficient: ")
qz=InputBox("Pressure Coefficient: [kPa]")
s=InputBox("Beam Load Width: [m]")
L=InputBox("Beam Span: [m]")

'Do Some Calculations
pn=Cpe*qz		'kPa
w=pn*s  		'kN/m
M=w*L^2/8		'kNm

'Summarise Inputs and Results in Report File
Wscript.Echo "Cpe = " & CStr(Cpe) 
Wscript.Echo "qz = " & CStr(qz) & " kPa"
Wscript.Echo "s = " & CStr(s) & " m"
Wscript.Echo "L = " & CStr(L) & " m"
Wscript.Echo "pn = " & FormatNumber(pn,2) & " kPa"
Wscript.Echo "w = " & FormatNumber(w,2) & " kN/m"
Wscript.Echo "M = " & FormatNumber(M,2) & " kNm"
