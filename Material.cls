VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Material"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public name As String
Public description As String
Public price As Double

Public quantity As Integer
Public quotQuantity As Integer
Public quantity_purch As Integer

Private is_Update As Boolean
Public min As Integer


Public Function getTotalPrice() As Double
    getTotalPrice = quantity_purch * price
End Function

Public Function updateQuantity() As Boolean
    updateQuantity = False
    If Not is_Update Then
        quantity = quantity + quantity_purch
        updateQuantity = True
    End If
End Function

Public Function isUpdate() As Boolean
    isUpdate = is_Update
End Function
