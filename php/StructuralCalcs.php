<?php
$Cpe=-0.7;
$qz=0.96;			//kPa
$pn=$Cpe*$qz;		//kPa
$w=$pn*3;			//kN/m
$L=6;				//m
$M=$w*pow($L,2)/8;	//kNm
echo "Moment: " . $M . " kNm";
?>
