(defun c:StructCalc ( / Cpe qz pn w L M)
	(setq 
		Cpe -0.7
		qz 0.96
		pn (* Cpe qz) 
		w (* pn 3) 
		L 6
		M  ( / (* w (* L L)) 8)
	)	
	(princ (strcat "Moment " (rtos M) " kNm") )
)
