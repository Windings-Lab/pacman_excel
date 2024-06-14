VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit
Private Sub Worksheet_Activate()
    If Mode = 0 Then
        Application.OnKey "{LEFT}", "Left"
        Application.OnKey "{RIGHT}", "Right"
        Application.OnKey "{UP}", "Up"
        Application.OnKey "{DOWN}", "Down"
    End If
End Sub
Private Sub Worksheet_Deactivate()
        Application.OnKey "{LEFT}"
        Application.OnKey "{RIGHT}"
        Application.OnKey "{UP}"
        Application.OnKey "{DOWN}"
End Sub


