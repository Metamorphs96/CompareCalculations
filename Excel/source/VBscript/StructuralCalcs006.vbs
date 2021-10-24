Option Explicit

Sub MainApplicationV6()
  Dim struct1(2)
  Dim i
  Dim k
  
  For i = 0 To UBound(struct1)
    k = i + 1
    Set struct1(i) = New clsStructure3d
    With struct1(i)
      .initialise
      Call .SetParameters(-0.7, 0.96, 3, 3 * k)
      .cprint
    End With
  Next

End Sub
