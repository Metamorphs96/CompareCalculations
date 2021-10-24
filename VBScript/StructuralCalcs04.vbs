Option Explicit

'Example to Testing piping data input and output results
' cscript structuralcalcs03.vbs < data.txt

Dim Cpe		'External Pressure Coefficient
Dim qz		'Site Reference Wind Pressure [kPa]
Dim s		'Beam Spacing = Load Width [m]
Dim L		'Beam Span [m]
Dim pn		'Design Pressure [kPa]
Dim w		'Uniformly Distributed Design Load [kN/m]
Dim M		'Bending Moment [kNm]

Dim StdIn, StdOut
Dim fpText
Dim fso

Set fso = CreateObject("Scripting.FileSystemObject")
set fpText = fso.CreateTextFile("results.txt", True)


Set StdIn = WScript.StdIn
Set StdOut = WScript.StdOut

'Get Values of Input Parameters
StdOut.WriteLine "External Surface Pressure Coefficient"
Cpe = StdIn.ReadLine

StdOut.WriteLine "Site Reference Pressure [kPa]"
qz = StdIn.ReadLine

StdOut.WriteLine "Load Width = Beam Spacing [m]"
s = StdIn.ReadLine

StdOut.WriteLine "Beam Span [m]"
L = StdIn.ReadLine

'Do Some Calculations
pn=Cpe*qz		'kPa
w=pn*s  		'kN/m
M=w*L^2/8		'kNm

'Summarise Inputs and Results in Report File
fpText.WriteLine "Cpe = " & CStr(Cpe) 
fpText.WriteLine "qz = " & CStr(qz) & " kPa"
fpText.WriteLine "s = " & CStr(s) & " m"
fpText.WriteLine "L = " & CStr(L) & " m"
fpText.WriteLine "pn = " & FormatNumber(pn,2) & " kPa"
fpText.WriteLine "w = " & FormatNumber(w,2) & " kN/m"
fpText.WriteLine "M = " & FormatNumber(M,2) & " kNm"

WScript.Echo "Results in File: results.txt"
WScript.Echo "All Done!"