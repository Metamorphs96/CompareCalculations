Option Explicit

'NB: Null is not an zero length string "", nor is an empty variable
Function Nz(x, valueifnull)
  Dim isNullValue 
  Dim vType
  
  'Debug.Print TypeName(x)
  isNullValue = IsNull(x)
  vType = VarType(x) 'NB: = 1, if null
  
  If isNullValue Then
    'Debug.Print "Has Null Value"
    'NB: Cannot check varType since varType=1 for Null
    'Need to check type of variable value being assigned to
    'To do so would appear to require another variable be passed
    'therefore may as well pass alternative value if null is found.
    'Need to find out how NZ built into access can allow valueifnull to be optional.
    'Can possibly get type from database field.
    'Until then change all access code to explicitly define the valieifnull variable
    
    Nz = valueifnull
  Else
    Nz = x
  End If

End Function
