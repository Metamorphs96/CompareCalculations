my $Cpe=-0.7;
my $qz=0.96;		#kPa
my $pn=$Cpe*$qz;	#kPa
my $w=$pn*3;		#kN/m
my $L=6;			#m
my $M=$w*$L**2/8;	#kNm
printf( "Moment: %.2f kNm" , $M );
