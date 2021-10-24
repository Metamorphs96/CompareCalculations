Attribute VB_Name = "StructuralCalcs006"
Option Explicit

Sub MainApplicationV6()
  Dim struct1(1) As clsStructure3d
  Dim i As Integer
  Dim k As Integer
  
  For i = 0 To 1
    k = i + 1
    Set struct1(i) = New clsStructure3d
    With struct1(i)
      .initialise
      Call .SetParameters(-0.7, 0.96, 3, 3 * k)
      .cprint
    End With
  Next i

End Sub
