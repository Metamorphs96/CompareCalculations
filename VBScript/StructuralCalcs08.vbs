Option Explicit

Sub StructuralCalcsV1(mdbDataPath)
	Const dbOpenDynaset = 2
	Dim dbe,db
	Dim StructuralData

	Dim Cpe 'External Pressure Coefficient
	Dim qz  'Site Reference Wind Pressure [kPa]
	Dim s   'Beam Spacing = Load Width [m]
	Dim L   'Beam Span [m]
	Dim pn  'Design Pressure [kPa]
	Dim w   'Uniformly Distributed Design Load [kN/m]
	Dim M   'Bending Moment [kNm]
	
	Set dbe = CreateObject("DAO.DBEngine.36")
	Set db = dbe.Workspaces(0).OpenDatabase(mdbDataPath, , True)
	Set StructuralData = db.OpenRecordset("StructuralCalcsV1", dbOpenDynaset)

	With StructuralData
		.MoveFirst
		Do
			Cpe = .Fields("Cpe").Value
			qz = .Fields("qz").Value
			s = .Fields("s").Value
			L = .Fields("L").Value

			pn = Cpe * qz         'kPa
			w = pn * s            'kN/m
			M = w * L ^ 2 / 8     'kNm

			WScript.Echo "Cpe = " & CStr(Cpe)
			WScript.Echo "qz = " & CStr(qz) & " kPa"
			WScript.Echo "s = " & CStr(s) & " m"
			WScript.Echo "L = " & CStr(L) & " m"
			WScript.Echo "pn = " & FormatNumber(pn, 2) & " kPa"
			WScript.Echo "w = " & FormatNumber(w, 2) & " kN/m"
			WScript.Echo "M = " & FormatNumber(M, 2) & " kNm"
			WScript.Echo "------------------------------------"
			.MoveNext
		Loop Until .EOF
	End With
 
End Sub

Sub StructuralCalcsV2(mdbDataPath)
	Const dbOpenDynaset = 2
	Dim dbe,db
	Dim StructuralData
	Dim StructuralResults


	Dim Cpe	'External Pressure Coefficient
	Dim qz	'Site Reference Wind Pressure [kPa]
	Dim s		'Beam Spacing = Load Width [m]
	Dim L		'Beam Span [m]
	Dim pn	'Design Pressure [kPa]
	Dim w		'Uniformly Distributed Design Load [kN/m]
	Dim M		'Bending Moment [kNm]

	Set dbe = CreateObject("DAO.DBEngine.36")
	Set db = dbe.Workspaces(0).OpenDatabase(mdbDataPath, , True)
	db.Execute "DELETE StructuralResults.* FROM StructuralResults;"
	Set StructuralData = db.OpenRecordset("StructuralCalcsV1", dbOpenDynaset)
	Set StructuralResults = db.OpenRecordset("StructuralResults", dbOpenDynaset)

	With StructuralData
		.MoveFirst
		Do
			Cpe = .Fields("Cpe").Value
			qz = .Fields("qz").Value
			s = .Fields("s").Value
			L = .Fields("L").Value

			pn = Cpe * qz         'kPa
			w = pn * s            'kN/m
			M = w * L ^ 2 / 8     'kNm

			StructuralResults.AddNew
			StructuralResults.Fields("pn").Value = pn
			StructuralResults.Fields("w").Value = w
			StructuralResults.Fields("M").Value = M
			StructuralResults.Update

			.MoveNext
		Loop Until .EOF
	End With

	WScript.Echo "All Done!"
 
End Sub


Sub cMain '(ByVal cmdArgs() )
	Dim fso, WshShell, objArgs
	
	'General
	Dim mdbDataPath
	Dim userDocPath
	
	
	
	WScript.Echo "Main ..."
	Set fso = CreateObject("Scripting.FileSystemObject")
	Set WshShell = CreateObject("WScript.Shell")
	userDocPath = WshShell.SpecialFolders("MyDocuments")
	Set objArgs = WScript.Arguments

	mdbDataPath = userDocPath & "\eCalcs\Library\Materials2\StructuralCalcs.mdb"

	StructuralCalcsV2 mdbDataPath

	WScript.Echo "... Main"
	WScript.Echo "All Done!"
	
End Sub
