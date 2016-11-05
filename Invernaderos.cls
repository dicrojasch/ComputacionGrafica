VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Invernaderos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private is_Radio As Boolean
Private is_AnchoLargo As Boolean
Private is_AreaLado As Boolean

Private radio As Double
Private ancho As Double
Private largo As Double
Private area As Double
Private lado As Double



Public Sub setRadio(radio_ As Double)
    radio = radio_
    is_Radio = True
    is_AnchoLargo = False
    is_AreaLado = False
End Sub


Public Sub setAnchoLargo(ancho_ As Double, largo_ As Double)
    ancho = ancho_
    largo = largo_
    is_Radio = False
    is_AnchoLargo = True
    is_AreaLado = False
End Sub

Public Sub setAreaLado(area_ As Double, lado_ As Double)
    area = area_
    lado = lado_
    is_Radio = False
    is_AnchoLargo = False
    is_AreaLado = True
End Sub

Public Function getRadio() As Double
    If isRadio Then
        getRadio = radio
    Else
        getRadio = -1
    End If
End Function

Public Function getAnchoLargo() As Variant() ' retorna un array de 2 elementos (ancho, largo)
    Dim answer() As Variant
    If isAnchoLargo Then
        ReDim answer(2)
        answer(1) = ancho
        answer(2) = largo
    Else
        ReDim answer(1)
        answer(1) = -1
    End If
    getAnchoLargo = answer
End Function

Public Function getAreaLado() As Variant ' retorna un array de 2 elementos (area, lado)
    Dim answer() As Variant
    If isAreaLado Then
        ReDim answer(2)
        answer(1) = area
        answer(2) = lado
    Else
        ReDim answer(1)
        answer(1) = -1
    End If
    getAreaLado = answer
End Function

Public Function isRadio() As Boolean
    isRadio = is_Radio
End Function

Public Function isAnchoLargo() As Boolean
    isAnchoLargo = is_AnchoLargo
End Function

Public Function isAreaLado() As Boolean
    isAreaLado = is_AreaLado
End Function


Public Sub JSONtoInvernaderos(objJson As Object)
    unidades = objJson.Item("medidas").Item("unidades")
    If isRadio Then radio = objJson.Item("medidas").Item("radio")
    If isAnchoLargo Then ancho = objJson.Item("medidas").Item("ancho")
                         largo = objJson.Item("medidas").Item("largo")
    
    If isAreaLado Then area = objJson.Item("medidas").Item("area")
                       lado = objJson.Item("medidas").Item("lado")
End Sub

Public Function toString() As String
    Dim temp() As Variant
    ReDim temp(2)
    If Me.isRadio Then
        toString = "Tipo : Radio ; " & "Radio = " & Me.getRadio & " ."
    ElseIf Me.isAnchoLargo Then
        temp = Me.getAnchoLargo()
        toString = "Tipo : Ancho y Largo ; " & "Ancho = " & temp(1) & " , " & "Largo = " & temp(2) & " ."
    ElseIf Me.isAreaLado Then
        temp = Me.getAreaLado()
        toString = "Tipo : Area y Lado ; " & "Area = " & temp(1) & " , " & "Lado = " & temp(2) & " ."
    
    Else
        MsgBox "Se debe Inicializar el objeto con la funcion 'setRadio' o 'setAnchoLargo' o 'setAreaLado'. "
    End If
End Function
