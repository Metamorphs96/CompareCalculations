Cpe=-0.7
qz=0.96			'kPa
s=3				'm
L=6				'm
pn=Cpe*qz		'kPa
w=pn*s  		'kN/m
M=w*L^2/8		'kNm
Wscript.Echo "Moment: ", FormatNumber(M,2),"kNm"

