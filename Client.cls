VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Client"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public id As Integer
Public firstName As String
Public lastname As String
Public email As String


Public Sub JSONtoClient(objJson As Object)
    If Not objetoJson Is Nothing Then
        email = Split(objetoJson.Item("datosUsuario").Item("email"), " ")(0)
        firstName = objetoJson.Item("datosUsuario").Item("nombre")
        lastname = objetoJson.Item("datosUsuario").Item("apellidos")
    End If
End Sub
