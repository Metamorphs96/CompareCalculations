Option Explicit

Sub MainApplicationV5()
  Dim struct1
  
  Set struct1 = New clsStructure3d
  
  With struct1
    .initialise
    Call .SetParameters(-0.7, 0.96, 3, 6)
    .cprint
  End With

End Sub



