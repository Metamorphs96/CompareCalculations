Option Explicit

Dim mdbDataPath

Sub setCharacteristic(characteristicName, characteristicValue)
	Const dbOpenDynaset = 2
	Dim dbe,db
	Dim tbl
	Dim StrCriteria

	Set dbe = CreateObject("DAO.DBEngine.36")
	Set db = dbe.Workspaces(0).OpenDatabase(mdbDataPath, , True)
	Set tbl = db.OpenRecordset("StructuralCalcsV3", dbOpenDynaset)

	StrCriteria = "Characteristic = " & """" & characteristicName & """"
	tbl.FindFirst StrCriteria
	If Not (tbl.NoMatch) Then
		tbl.Edit
		tbl.Fields("Value").Value = characteristicValue
		tbl.Update
	End If
	tbl.Close
	
End Sub



Function getCharacteristic(characteristicName)
	Const dbOpenDynaset = 2
	Dim dbe,db
	Dim tbl
	Dim StrCriteria

	Set dbe = CreateObject("DAO.DBEngine.36")
	Set db = dbe.Workspaces(0).OpenDatabase(mdbDataPath, , True)
	Set tbl = db.OpenRecordset("StructuralCalcsV3", dbOpenDynaset)

	StrCriteria = "Characteristic = " & """" & characteristicName & """"
	tbl.FindFirst StrCriteria
	If Not (tbl.NoMatch) Then
		getCharacteristic = Nz(tbl.Fields("Value").Value,"")
	Else
		getCharacteristic = ""
	End If
	tbl.Close
	
End Function


Sub StructuralCalcsV3()

	Dim Cpe	'External Pressure Coefficient
	Dim qz	'Site Reference Wind Pressure [kPa]
	Dim s		'Beam Spacing = Load Width [m]
	Dim L		'Beam Span [m]
	Dim pn	'Design Pressure [kPa]
	Dim w		'Uniformly Distributed Design Load [kN/m]
	Dim M		'Bending Moment [kNm]


	Cpe = getCharacteristic("Cpe")
	qz = getCharacteristic("qz")
	s = getCharacteristic("s")
	L = getCharacteristic("L")

	pn = Cpe * qz         'kPa
	w = pn * s            'kN/m
	M = w * L ^ 2 / 8     'kNm

	Call setCharacteristic("pn", CStr(pn))
	Call setCharacteristic("w", CStr(w))
	Call setCharacteristic("M", CStr(M))

	WScript.Echo "All Done!"

End Sub


Sub cMain '(ByVal cmdArgs() )
	Dim fso, WshShell, objArgs
	
	'General
	Dim userDocPath
	
	
	
	WScript.Echo "Main ..."
	Set fso = CreateObject("Scripting.FileSystemObject")
	Set WshShell = CreateObject("WScript.Shell")
	userDocPath = WshShell.SpecialFolders("MyDocuments")
	Set objArgs = WScript.Arguments

	mdbDataPath = userDocPath & "\eCalcs\Library\Materials2\StructuralCalcs.mdb"

	StructuralCalcsV3

	WScript.Echo "... Main"
	WScript.Echo "All Done!"
	
End Sub
