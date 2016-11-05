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
    getPrice = 0
    If producto.price > 0 Then
        getPrice = producto.price * (1 + benefit)
    End If
    
End Function




Public Sub JSONtoInvernaderos(objJson As Object)
    
    unidades = objJson.Item("medidas").Item("unidades")
    If isRadio Then radio = objJson.Item("medidas").Item("radio")
    If isAnchoLargo Then ancho = objJson.Item("medidas").Item("ancho")
                         largo = objJson.Item("medidas").Item("largo")
    
    If isAreaLado Then area = objJson.Item("medidas").Item("area")
                       lado = objJson.Item("medidas").Item("lado")
    
End Sub
