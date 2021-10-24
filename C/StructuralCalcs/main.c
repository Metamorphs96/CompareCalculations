#include <stdio.h>
#include <math.h>

int main()
{
	float Cpe,qz,pn,w,L,M;

	Cpe=-0.7;
	qz=0.96;			//kPa
	pn=Cpe*qz;			//kPa
	w=pn*3.000;			//kN/m
	L=6.000;			//m
	M=w*pow(L,2)/8;		//kNm

	printf("Moment: %5.2f" , M);
}

