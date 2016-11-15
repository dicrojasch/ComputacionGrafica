VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Invernaderos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Public tipo As String
Public ancho As Double
Public largo As Double
Public alto As Double

Public Sub JSONtoInvernaderos(objJson As Object)
    tipo = objJson.item("medidas").item("tipo")   'para cambiar...
    ancho = objJson.item("medidas").item("ancho") 'para cambiar...
    largo = objJson.item("medidas").item("largo") 'para cambiar...
    alto = objJson.item("medidas").item("alto")   'para cambiar...
End Sub

Public Function toString() As String
    toString = "Tipo: " & tipo & vbCrLf
    toString = toString & "Ancho: " & ancho & vbCrLf
    toString = toString & "Largo: " & largo & vbCrLf
    toString = toString & "Alto: " & alto & vbCrLf
End Function



