VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Product"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private name As String
Public id As Integer
Private description As String
Public price As Double
Private isFormaleta As Boolean
Private isInvernadero As Boolean
Private isMaquinaria As Boolean
Private materials As New Dictionary




Public Function setFormaleta(formaleta As Formaletas)
    If Not formaleta Is Nothing Then
        name = "Formaleta"
        description = formaleta.toString
        isFormaleta = True
        isInvernadero = False
        isMaquinaria = False
    End If
End Function

Public Function setInvernadero(Invern As Invernaderos)
    If Not Invern Is Nothing Then
        name = "Invernadero"
        description = Invern.toString
        isFormaleta = False
        isInvernadero = True
        isMaquinaria = False
    End If
End Function

Public Function getDescription() As String
    getDescription = description
End Function

Public Function getName() As String
    getName = name
End Function

Public Function existsMaterials() As Boolean
    If materials.Count > 0 Then
        existsMaterials = True
    Else
        existsMaterials = False
    End If
End Function

Public Function getMaterial(name As String) As Material
    If materials.Exists(name) Then
        getMaterial = materials(name)
    End If
End Function

Public Function getMaterialSize() As Integer
    getMaterialSize = materials.Count
End Function

Public Function getAllMaterials() As Collection
    Set getAllMaterials = New Collection
    For Each Item In materials.Items
        getAllMaterials.Add Item
    Next Item
End Function

Public Function addMaterial(mat As Material) As Boolean
    If Not materials.Exists(mat.name) Then
        Call materials.Add(mat.name, mat)
        addMaterial = True
    Else
        addMaterial = False
    End If
End Function

Public Function deleteMaterial(mat As Material) As Boolean
    If materials.Exists(mat.name) Then
        materials.Remove mat.name
    End If
    deleteMaterial = True
End Function


Public Sub JSONtoProduct(objJson As Object)
    If Not objetoJson Is Nothing Then
        email = Split(objetoJson.Item("datosUsuario").Item("email"), " ")(0)
        firstName = objetoJson.Item("datosUsuario").Item("nombre")
        lastname = objetoJson.Item("datosUsuario").Item("apellidos")
    End If
End Sub

Public Function mailDescription() As String
    If Not isFormaleta And Not isInvernadero And Not isMaquinaria Then
        Debug.Print "Error, no se ha inicializado el producto "
    Else
        mailDescription = "Producto: " & name & vbCrLf & "Descripcion: " & vbCrLf & description
    End If
End Function

Public Function toString() As String
    If Not isFormaleta And Not isInvernadero And Not isMaquinaria Then
        MsgBox "Se debe Inicializar el objeto con la funcion ''. "
    Else
        toString = "Producto: " & name & vbCrLf & "Descripcion: " & vbCrLf & description & "."
        If price <> 0 Then
            toString = toString & " , price = " & price
        End If
        toString = toString & "; Materiales : "
        For Each Item In materials.Items
            toString = toString & ", " & Item
        Next Item
    End If
End Function

