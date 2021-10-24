Cpe=-0.7;
qz=0.96;				//kPa
pn=Cpe*qz;				//kPa
w=pn*3.000;  			//kN/m
L=6.000;				//m
M=w*Math.pow(L,2)/8;	//kNm

WScript.Echo("Moment: " + M.toFixed(2));

