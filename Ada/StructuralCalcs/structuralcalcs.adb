with 
	ada.Text_IO,
	ada.long_float_text_io;


procedure StructuralCalcs is

  Cpe : long_float;
  qz : long_float;
  pn : long_float;
  s : long_float;
  w : long_float;
  L : long_float;
  M : long_float;

begin

	Cpe:=-0.7;
	qz:=0.96;		--kPa
  	pn:=Cpe*qz;		--kPa
   	s:=3.0;			--m
	w:=pn*s;  		--kN/m
	L:=6.0;			--m
	M:=w*L*L/8.0;		--kNm
	Ada.Text_IO.Put("Moment: ");
	Ada.long_float_text_io.Put(M,5,2,0);
	Ada.Text_IO.Put(" kNm");


end StructuralCalcs;
