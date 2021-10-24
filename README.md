# CompareCalculations
Purpose of the exercise is to compare a simple common structural calculation sequence using available calculation tools, and compare to Excel/VBA and the VBE immediate window. The calculation sequence is:
```
  Cpe=-0.7  
  qz=0.96   'kPa  
  s=3       'm  
  L=6       'm  
  pn=Cpe*qz 'kPa  
  w=pn*s    'kN/m  
  M=w*L^2/8 'kNm
```
as written in VBscript. The other issue is the ease with which a report can be produced which provides the input and the output, as well as document the process between the input and output assuming its important to describe such. In many situations it is not necessary to provide detail as to how got from the inputs to the outputs, as it is an independent reviewers task to review the resultant specifications and produce their own calculations to confirm or refute the asserted suitability of a proposal.

The calculation sequence is simple but it can be extended. For example I have a multitude of VBA functions/procedures to calculate Cpe for various situtations, and similarly qz as to be calculated based on the conditions of the site. So whilst I have just given them values they are actually dependent on other inputs, as for the action-effect (M) its calculaation is highly variable and one of the biggest bottleneck to automating calculations. Structural design can be divided into following sequence of tasks:

1) Determine Dimension and Geometry
2) Determine Design-Actions
3) Determine Design-Action-Effects
4) Determine Member Sizes
5) Design Connection
6) Design Footings

Just about everything can be done in MS Excel/VBA except task (3), where depending on the structural form the calculation of design-action-effects becomes cumbersome or impractical: that is typically use external structural analysis software, and then have a problem interfacing with its required inputs and getting its outputs for further calculation. And python/jupyter notebooks whilst seemingly popular are not the solution to the problem of integration of applications and calculations.
