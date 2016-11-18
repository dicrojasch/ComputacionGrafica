VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Formaletas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public altura As Double
Public diamInterno As Double
Public AltRanura As Double
Public unidades As String

Public cPlate0 As String
Public cPlate90 As String
Public cPlate180 As String
Public cPlate270 As String
Public aFPlate0 As String
Public aFPlate45 As String
Public aFPlate90 As String
Public aFPlate135 As String
Public aFPlate180 As String
Public aFPlate225 As String
Public aFPlate270 As String
Public aFPlate315 As String
Public rVar0_90 As Boolean
Public rVar90_180 As Boolean
Public rVar180_270 As Boolean
Public rVar270_0 As Boolean
Private isInit As Boolean

Public Sub InitFormaletas(altura_ As Double, diamInterno_ As Double, AltRanura_ As Double)

    altura = altura_
    diamInterno = diamInterno_
    AltRanura = AltRanura_
    unidades = ""
    cPlate0 = ""
    cPlate90 = ""
    cPlate180 = ""
    cPlate270 = ""
    aFPlate0 = ""
    aFPlate45 = ""
    aFPlate90 = ""
    aFPlate135 = ""
    aFPlate180 = ""
    aFPlate225 = ""
    aFPlate270 = ""
    aFPlate315 = ""
    rVar0_90 = False
    rVar90_180 = False
    rVar180_270 = False
    rVar270_0 = False
    isInit = True
    
End Sub


Private Function getValueString(objJson As Object, value As String) As String
    getValueString = "N/A"
    If objJson.Exists(value) Then
        If objJson.item(value).item("activado") Then
            getValueString = objJson.item(value).item("texto")
        End If
    End If
End Function

Private Function getValueBool(objJson As Object, value As String) As String
    getValueBool = False
    If objJson.Exists(value) Then
        If objJson.item(value).item("activado") Then
            getValueBool = objJson.item(value).item("activado")
        End If
    End If
End Function

Public Sub JSONtoFormaleta(objJson As Object)
        
    unidades = objJson.item("medidas").item("unidades")
    altura = objJson.item("medidas").item("altura")
    diamInterno = objJson.item("medidas").item("diametroInterno")
    AltRanura = objJson.item("medidas").item("alturaRanura")
    
    cPlate0 = getValueString(objJson.item("opciones"), "CP_0")
    cPlate90 = getValueString(objJson.item("opciones"), "CP_90")
    cPlate180 = getValueString(objJson.item("opciones"), "CP_180")
    cPlate270 = getValueString(objJson.item("opciones"), "CP_270")
    aFPlate0 = getValueString(objJson.item("opciones"), "AFP_0")
    aFPlate45 = getValueString(objJson.item("opciones"), "AFP_45")
    aFPlate90 = getValueString(objJson.item("opciones"), "AFP_90")
    aFPlate135 = getValueString(objJson.item("opciones"), "AFP_135")
    aFPlate180 = getValueString(objJson.item("opciones"), "AFP_180")
    aFPlate225 = getValueString(objJson.item("opciones"), "AFP_225")
    aFPlate270 = getValueString(objJson.item("opciones"), "AFP_270")
    aFPlate315 = getValueString(objJson.item("opciones"), "AFP_315")
    
    rVar0_90 = getValueBool(objJson.item("opciones"), "RV_0_90")
    rVar90_180 = getValueBool(objJson.item("opciones"), "RV_90_180")
    rVar180_270 = getValueBool(objJson.item("opciones"), "RV_180_270")
    rVar270_0 = getValueBool(objJson.item("opciones"), "RV_270_0")
    isInit = True
    
End Sub

Public Function toString() As String
    If isInit Then
        toString = "Altura = " & altura & vbCrLf & "Diametro Interno = " & diamInterno & vbCrLf & "Altura ranura = " & AltRanura & vbCrLf
             If cPlate0 <> "" Then
            toString = toString & "Cover Plate a 0 = " & cPlate0 & vbCrLf
        End If
        If cPlate90 <> "" Then
            toString = toString & "Cover Plate a 90 = " & cPlate90 & vbCrLf
        End If
        If cPlate180 <> "" Then
            toString = toString & "Cover Plate a 180 = " & cPlate180 & vbCrLf
        End If
        If cPlate270 <> "" Then
            toString = toString & "Cover Plate a 270 = " & cPlate270 & vbCrLf
        End If
        If aFPlate0 <> "" Then
            toString = toString & "Auto Filler Plate a 0 = " & aFPlate0 & vbCrLf
        End If
        If aFPlate45 <> "" Then
            toString = toString & "Auto Filler Plate a 45 = " & aFPlate45 & vbCrLf
        End If
        If aFPlate90 <> "" Then
            toString = toString & "Auto Filler Plate a 90 = " & aFPlate90 & vbCrLf
        End If
        If aFPlate135 <> "" Then
            toString = toString & "Auto Filler Plate a 135 = " & aFPlate135 & vbCrLf
        End If
        If aFPlate180 <> "" Then
            toString = toString & "Auto Filler Plate a 180 = " & aFPlate180 & vbCrLf
        End If
        If aFPlate225 <> "" Then
            toString = toString & "Auto Filler Plate a 225 = " & aFPlate225 & vbCrLf
        End If
        If aFPlate270 <> "" Then
            toString = toString & "Auto Filler Plate a 270 = " & aFPlate270 & vbCrLf
        End If
        If aFPlate315 <> "" Then
            toString = toString & "Auto Filler Plate a 315 = " & aFPlate315 & vbCrLf
        End If
        If rVar0_90 Then
            toString = toString & "Ranuras de Varillas entre 0 y 90" & vbCrLf
        End If
        If rVar90_180 Then
            toString = toString & "Ranuras de Varillas entre 90 y 180" & vbCrLf
        End If
        If rVar180_270 Then
            toString = toString & "Ranuras de Varillas entre 180 y 270" & vbCrLf
        End If
        If rVar270_0 Then
            toString = toString & "Ranuras de Varillas entre 270 y 0" & vbCrLf
        End If
    Else
        MsgBox "se debe inicializar el objeto con la funcion 'InitFormaletas'."
    End If
End Function
