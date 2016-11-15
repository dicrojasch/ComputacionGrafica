VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Purchases"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public id As Integer
Public provider_name As String
Private materials As New Dictionary
Public description As String
Private total_price As Double

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

Public Function getAllMaterials() As Collection
    Set getAllMaterials = New Collection
    For Each item In materials.Items
        getAllMaterials.Add item
    Next item
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

Public Function getTotalPrice() As Double
    getTotalPrice = 0
    For Each mat In materials.Items
        getTotalPrice = getTotalPrice + mat.getTotalPrice()
    Next mat
End Function

