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
Public state As Integer

Public Function getPrice() As Double
    getPrice = 0
    If producto.price > 0 Then
        getPrice = producto.price * (1 + benefit)
    End If
    
End Function




Public Sub JSONtoInvernaderos(objJson As Object)
    
    unidades = objJson.item("medidas").item("unidades")
    If isRadio Then radio = objJson.item("medidas").item("radio")
    If isAnchoLargo Then ancho = objJson.item("medidas").item("ancho")
                         largo = objJson.item("medidas").item("largo")
    
    If isAreaLado Then area = objJson.item("medidas").item("area")
                       lado = objJson.item("medidas").item("lado")
    
End Sub
