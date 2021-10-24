      PROGRAM StructuralCalc
      REAL Cpe,qz,pn,w,L,M
      Cpe=-0.7
      qz=0.96
      pn=Cpe*qz
      w=pn*3
      L=6
      M=w*L**2/8
      PRINT '(A)','Moment: ', M
	  STOP
      END

