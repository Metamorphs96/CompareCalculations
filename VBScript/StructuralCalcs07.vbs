Option Explicit

'Example structural calculations using hardcoded input/output filenames

Const ForReading = 1, ForWriting = 2, ForAppending = 8
Const TristateUseDefault = 2, TristateTrue = -1, TristateFalse = 0

Sub StructuralCalcs(dataFileName,resultFileName)
	Dim Cpe		'External Pressure Coefficient
	Dim qz		'Site Reference Wind Pressure [kPa]
	Dim s		'Beam Spacing = Load Width [m]
	Dim L		'Beam Span [m]
	Dim pn		'Design Pressure [kPa]
	Dim w		'Uniformly Distributed Design Load [kN/m]
	Dim M		'Bending Moment [kNm]

	Dim StdIn, StdOut
	Dim fpRpt, fpData
	Dim fso

	Set fso = CreateObject("Scripting.FileSystemObject")
	set fpRpt = fso.CreateTextFile(resultFileName, True)
	set fpData = fso.OpenTextFile(dataFileName)

	Set StdIn = WScript.StdIn
	Set StdOut = WScript.StdOut

	'Get Values of Input Parameters
	Cpe = fpData.ReadLine
	qz = fpData.ReadLine
	s = fpData.ReadLine
	L = fpData.ReadLine

	'Do Some Calculations
	pn=Cpe*qz		'kPa
	w=pn*s  		'kN/m
	M=w*L^2/8		'kNm

	'Summarise Inputs and Results in Report File
	fpRpt.WriteLine "Cpe = " & CStr(Cpe) 
	fpRpt.WriteLine "qz = " & CStr(qz) & " kPa"
	fpRpt.WriteLine "s = " & CStr(s) & " m"
	fpRpt.WriteLine "L = " & CStr(L) & " m"
	fpRpt.WriteLine "pn = " & FormatNumber(pn,2) & " kPa"
	fpRpt.WriteLine "w = " & FormatNumber(w,2) & " kN/m"
	fpRpt.WriteLine "M = " & FormatNumber(M,2) & " kNm"

	WScript.Echo "Results in File: <" & fso.GetFileName(resultFileName) & ">"

End Sub


Sub cMain '(ByVal cmdArgs() )
	Dim fso, WshShell, objArgs
	
	'General
	Dim fPath0, fPath1
	Dim fDrv , fPath, fName, fExt
	Dim ifullName 
	Dim ofullName 

	
	WScript.Echo "Main ..."
	Set fso = CreateObject("Scripting.FileSystemObject")
	Set WshShell = CreateObject("WScript.Shell")
	Set objArgs = WScript.Arguments
	
	
	' See if there are any arguments.
	If objArgs.Count = 1  Then
		fPath0 = objArgs(0)
		fPath1 = fso.GetAbsolutePathName(fPath0)
		
		
		fDrv = fso.GetDriveName(fPath1)
		fPath = fso.GetParentFolderName(fPath1)
		fName = fso.GetBaseName(fPath1)
		fExt = fso.GetExtensionName(fPath1) 
		
		' WScript.Echo fDrv
		' WScript.Echo fPath
		' WScript.Echo fso.GetBaseName(fPath1)
		' WScript.Echo fso.GetFileName(fPath1)
		' WScript.Echo fExt
		
		ifullName = fPath1
		ofullName = fPath & "\" & fName & ".rpt"
		
		WScript.Echo "INPUT: Data File: <" & fso.GetFileName(ifullName) & ">"
		WScript.Echo "OUTPUT: Data File: <" & fso.GetFileName(ofullName) & ">"
		
		StructuralCalcs ifullName,ofullName
		
		
	Else
		WScript.Echo "Not enough parameters: provide data file name, include file extension"
	End If
	
	WScript.Echo "... Main"
	WScript.Echo "All Done!"
	
End Sub
