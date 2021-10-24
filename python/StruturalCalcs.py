Cpe=-0.7
qz=0.96			#kPa
pn=Cpe*qz		#kPa
w=pn*3  		#kN/m
L=6			#m
M=w*L**2/8	        #kNm
print('Moment: %.2f'%(M))
print('All Done!')
