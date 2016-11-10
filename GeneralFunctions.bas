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

Function getMonth() As String
    Dim my_month As String
    Select Case Month(Date)
        Case 1
            getMonth = "Enero"
        Case 2
            getMonth = "Febrero"
        Case 3
            getMonth = "Marzo"
        Case 4
            getMonth = "Abril"
        Case 5
            getMonth = "Mayo"
        Case 6
            getMonth = "Junio"
        Case 7
            getMonth = "Julio"
        Case 8
            getMonth = "Agosto"
        Case 9
            getMonth = "Septiembre"
        Case 10
            getMonth = "Octubre"
        Case 11
            getMonth = "Noviembre"
        Case Else
            getMonth = "Diciembre"
    End Select
End Function


' Funcion para generar una fecha con el formato, <<dia>> del mes de <<mes>> del <<anio>>'
Function getDate() As String
    getDate = Day(Date) & " de " & getMonth & " del " & Year(Date)
End Function

Function parseJSON(strJson As String) As Object
    Dim clsJson As json
    Set clsJson = New json
    Set parseJSON = clsJson.parse(strJson)
End Function

Function FileExists(ByVal FileToTest As String) As Boolean
   FileExists = (Dir(FileToTest) <> "")
End Function

Sub DeleteFile(ByVal FileToDelete As String)
   If FileExists(FileToDelete) Then
      SetAttr FileToDelete, vbNormal
      Kill FileToDelete
   End If
End Sub

Sub createDirectory(directoryPath)
    directoryPath = path & directoryPath
    If Dir(directoryPath, vbDirectory) = "" Then
        MkDir directoryPath
    End If
End Sub



