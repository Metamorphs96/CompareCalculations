Option Explicit

'Vbscript Doesn't support Type...End Type 
'So implement record type as a class
Class Structure3D
  Public Cpe
  Public qz
  Public s
  Public L
  Public pn
  Public w
  Public M
End Class

Function getBendingMomentV2(ByRef struct1)
  
  With struct1
    .pn = .Cpe * .qz     'kPa
    .w = .pn * .s        'kN/m
    getBendingMomentV2 = .w * .L ^ 2 / 8 'kNm
  End With
  
End Function


Sub MainApplicationV4()
  Dim struct1
  
  Set struct1 = New Structure3D
  
  With struct1
    .Cpe = -0.7
    .qz = 0.96
    .s = 3
    .L = 6
    .M = getBendingMomentV2(struct1)
    
    WScript.Echo "Cpe = " & CStr(.Cpe)
    WScript.Echo "qz = " & CStr(.qz) & " kPa"
    WScript.Echo "s = " & CStr(.s) & " m"
    WScript.Echo "L = " & CStr(.L) & " m"
    WScript.Echo "pn = " & CStr(.pn) & " kPa"
    WScript.Echo "w = " & CStr(.w) & " kN/m"
    
    WScript.Echo "M = " & FormatNumber(.M, 2) & " kNm"
    WScript.Echo "------------------------------------"
  End With
  
End Sub


