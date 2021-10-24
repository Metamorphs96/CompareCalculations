Option Explicit

Dim mdbDataPath

Sub adoStructuralCalcsV1()
  Dim dbCon
  Dim StructuralData
  Dim connStr
  
  Dim Cpe 'External Pressure Coefficient
  Dim qz  'Site Reference Wind Pressure [kPa]
  Dim s   'Beam Spacing = Load Width [m]
  Dim L   'Beam Span [m]
  Dim pn  'Design Pressure [kPa]
  Dim w   'Uniformly Distributed Design Load [kN/m]
  Dim M   'Bending Moment [kNm]
  
  Set dbCon = CreateObject( "ADODB.Connection" )
  'dbCon.Provider = "Microsoft.Jet.OLEDB.4.0"
  'dbCon.Provider = "Microsoft.ACE.OLEDB.12.0"
  
  connStr = "Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" & mdbDataPath & ";"
  'connStr = "“Driver={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=" & mdbDataPath & ";"
  WScript.Echo connStr
  dbCon.Open connStr
  

  set StructuralData=CreateObject("ADODB.recordset")
  StructuralData.Open "Select * From StructuralCalcsV1", dbCon

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
 
  WScript.Echo "All Done!"
End Sub

Sub adoStructuralCalcsV2()
  Const adOpenDynamic=2
  Const adLockOptimistic=3
  Dim dbCon
  Dim StructuralData
  Dim StructuralResults
  
  Dim Cpe 'External Pressure Coefficient
  Dim qz  'Site Reference Wind Pressure [kPa]
  Dim s   'Beam Spacing = Load Width [m]
  Dim L   'Beam Span [m]
  Dim pn  'Design Pressure [kPa]
  Dim w   'Uniformly Distributed Design Load [kN/m]
  Dim M   'Bending Moment [kNm]
  
  Set dbCon = CreateObject( "ADODB.Connection" )
  'dbCon.Provider = "Microsoft.Jet.OLEDB.4.0"
  dbCon.Provider = "Microsoft.ACE.OLEDB.12.0"
  dbCon.Open mdbDataPath
  
  set StructuralData=CreateObject("ADODB.recordset")
  StructuralData.Open "Select * From StructuralCalcsV1", dbCon

  set StructuralResults=CreateObject("ADODB.recordset")
  StructuralResults.Open "Select * From StructuralResults", dbCon, adOpenDynamic, adLockOptimistic
  
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
	Dim userDocPath
	
	WScript.Echo WScript.FullName
	
	WScript.Echo "Main ..."
	Set fso = CreateObject("Scripting.FileSystemObject")
	Set WshShell = CreateObject("WScript.Shell")
	userDocPath = WshShell.SpecialFolders("MyDocuments")
	Set objArgs = WScript.Arguments

	mdbDataPath = userDocPath & "\eCalcs\Library\Materials2\StructuralCalcs.accdb"
	'mdbDataPath = "C:\Users2\AppDev\Materials2\StructuralCalcs.accdb"
	
	WScript.Echo mdbDataPath
	
	adoStructuralCalcsV1

	WScript.Echo "... Main"
	WScript.Echo "All Done!"
	
End Sub
