program StructuralCalc;

Var
  Cpe : double;
  qz : double;
  pn : double;
  w : double;
  L : double;
  M : double;

begin

Cpe:=-0.7;
qz:=0.96;		{kPa}
pn:=Cpe*qz;		{kPa}
w:=pn*3;  		{kN/m}
L:=6;			{m}
M:=w*L*L/8;		{kNm}
writeln('Moment: ', M:4:2);

end.

