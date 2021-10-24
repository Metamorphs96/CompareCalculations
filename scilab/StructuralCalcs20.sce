clear,clc,clf;

Cpe=1.2
qz=[0.69,0.96,1.5]			//kPa
pn=Cpe*qz		//kPa
s=3             //m
w=pn*s			//kN/m
L=6				//m
M=(L^2/8)*w		//kNm
L=0:1:11
M=zeros(3,12)
printf('Start Loop\n')
for i=1:3
    for j=1:1:12
        printf('%d %d %8.3f\n',i,j,L(j))
        M(i,j)=w(i)*L(j)^2/8
    end
end
plot(L,M)
title('Moment vs Span','fontsize',4);
xlabel('Span [m]')
ylabel('Moment [kNm]')
