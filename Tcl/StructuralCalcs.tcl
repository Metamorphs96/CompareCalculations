set Cpe -0.7
set qz 0.96
set pn [expr {$Cpe*$qz}]
set w [expr {$pn*3.0}]
set L 6.0	
set M [expr {$w*$L*$L/8.0}]
puts "Moment: $M kNm"
