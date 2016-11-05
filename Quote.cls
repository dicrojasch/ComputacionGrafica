VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Quote"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public cliente As Client
Public producto As Product
Public time_response As Long 'milisegundos
Public benefit As Double

Public Function getPrice() As Double
    getPrice = producto.price * (1 + benefit)
End Function




