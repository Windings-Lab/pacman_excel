VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisBook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub Workbook_Open()
    Application.OnKey "{LEFT}", "Left"
    Application.OnKey "{RIGHT}", "Right"
    Application.OnKey "{UP}", "Up"
    Application.OnKey "{DOWN}", "Down"
    Begin
End Sub