VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PhysicalCharacteristic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public key As Integer
Public Name As String
Public Value As Double 'String or variant may be more suitable for general case
Public Units As String
Public formatStr As String

Sub initialise()
  key = 0
  Name = ""
  Value = 0
  Units = ""
  formatStr = "0.00"
End Sub

Sub setCharacteristicKeyed(xkey As Integer, _
            xName As String, xValue As Double, xUnits As String, _
            xFormatStr As String)
  
  key = xkey
  Name = xName
  Value = xValue
  Units = xUnits
  formatStr = xFormatStr
  
End Sub

Sub setCharacteristic(xName As String, xValue As Double, xUnits As String, xFormatStr As String)
  
  Name = xName
  Value = xValue
  Units = xUnits
  formatStr = xFormatStr
  
End Sub

Sub setCharacteristicDefinition(xName As String, xUnits As String, xFormatStr As String)
  
  Name = xName
  Units = xUnits
  formatStr = xFormatStr
  
End Sub


Function sprint() As String
  sprint = Name & ": " & Format(Value, formatStr) & " " & Units
End Function

Sub cprint()
  Debug.Print sprint
End Sub
