Attribute VB_Name = "GeneralFunctions"
Public Function Pause(NumberOfSeconds As Variant)
    On Error GoTo Error_GoTo

    Dim PauseTime As Variant
    Dim Start As Variant
    Dim Elapsed As Variant

    PauseTime = NumberOfSeconds
    Start = Timer
    Elapsed = 0
    Do While Timer < Start + PauseTime
        Elapsed = Elapsed + 1
        If Timer = 0 Then
            ' Crossing midnight
            PauseTime = PauseTime - Elapsed
            Start = 0
            Elapsed = 0
        End If
        DoEvents
    Loop

Exit_GoTo:
    On Error GoTo 0
    Exit Function
Error_GoTo:
    Debug.Print Err.Number, Err.description, Erl
    GoTo Exit_GoTo
End Function

' Funcion para generar una fecha con el formato, <<dia>> del mes de <<mes>> del <<anio>>'
Function getDate() As String
    Dim my_month As String
    Select Case Month(Date)
        Case 1
            my_month = "Enero"
        Case 2
            my_month = "Febrero"
        Case 3
            my_month = "Marzo"
        Case 4
            my_month = "Abril"
        Case 5
            my_month = "Mayo"
        Case 6
            my_month = "Junio"
        Case 7
            my_month = "Julio"
        Case 8
            my_month = "Agosto"
        Case 9
            my_month = "Septiembre"
        Case 10
            my_month = "Octubre"
        Case 11
            my_month = "Noviembre"
        Case Else
            my_month = "Diciembre"
    End Select
    getDate = Day(Date) & " de " & my_month & " del " & Year(Date)
End Function

Function parseJSON(strJson As String) As Object
    Dim clsJson As json
    Set clsJson = New json
    Set parseJSON = clsJson.parse(strJson)
End Function

Sub DeleteFile(ByVal FileToDelete As String)
   If FileExists(FileToDelete) Then
      SetAttr FileToDelete, vbNormal
      Kill FileToDelete
   End If
End Sub
