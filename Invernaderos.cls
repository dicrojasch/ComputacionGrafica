VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Invernaderos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Invernaderos

Public tipo As String
Public ancho As Double
Public largo As Double
Public alto As Double

Public Sub JSONtoInvernaderos(objJson As Object)
    tipo = objJson.Item("medidas").Item("tipo")   'para cambiar...
    ancho = objJson.Item("medidas").Item("ancho") 'para cambiar...
    largo = objJson.Item("medidas").Item("largo") 'para cambiar...
    alto = objJson.Item("medidas").Item("alto")   'para cambiar...
End Sub

Public Function toString() As String
    toString = "Tipo: " & tipo & vbCrLf
    toString = toString & "Ancho: " & ancho & vbCrLf
    toString = toString & "Largo: " & largo & vbCrLf
    toString = toString & "Alto: " & alto & vbCrLf
End Function


