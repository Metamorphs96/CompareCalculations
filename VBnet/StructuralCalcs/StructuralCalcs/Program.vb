'
' Created by SharpDevelop.
' User: Conrad
' Date: 6/10/2018
' Time: 11:37
' 
' To change this template use Tools | Options | Coding | Edit Standard Headers.
'
Module Program
	Sub Main()	
		Dim Cpe As Double
		Dim qz As Double
		Dim pn As Double 
		Dim s As Double
		Dim w As Double 
		Dim L As Double
		Dim M As Double			
		
		Cpe = -0.7
		qz = 0.96   		'kPa
		pn = Cpe * qz 		'kPa
		s = 3				'm
		w = pn * s  		'kN/m
		L = 6       		'm
		M = w * L ^ 2 / 8 	'kNm
		
		Console.WriteLine("Moment: " & Format(M, "0.00"))

	End Sub
End Module
