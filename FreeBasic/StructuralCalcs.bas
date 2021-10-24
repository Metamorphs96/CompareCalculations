#include once "string.bi" 'required for format function

Dim Cpe as Double
Dim qz as Double
Dim pn as Double
Dim w as Double
Dim L as Double
Dim M as Double

Cpe=-0.7
qz=0.96			'kPa
pn=Cpe*qz		'kPa
w=pn*3  		'kN/m
L=6				'm
M=w*L^2/8		'kNm
Print "Moment: ", Format(M,"0.00")



