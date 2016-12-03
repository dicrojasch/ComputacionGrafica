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
    Dim folders() As String
    folders = Split(directoryPath, "\")
    Dim cumulative As String
    cumulative = ""
    For Each Folder In folders
        cumulative = cumulative & Folder & "\"
        If Dir(cumulative, vbDirectory) = "" Then
            MkDir cumulative
        Else
            Debug.Print "La Carpeta " & cumulative & " ya existe."
        End If
    Next Folder
End Sub
Sub actualizaPrecios()
    Dim ExcelApp As Excel.Application
    Set ExcelApp = New Excel.Application
    Dim database As New GraficaDB
    Call database.ConnectDB(DBServer, schema, user, password)
    ExcelApp.Workbooks.Open (pathInvernaderos & "Precios.xlsx")
    With ExcelApp.ActiveSheet
        .Cells(2, 3).value = database.getMaterial("precioTC2218").price
        .Cells(3, 3).value = database.getMaterial("precioT10133").price
        .Cells(4, 3).value = database.getMaterial("precioBR141").price
        .Cells(5, 3).value = database.getMaterial("precioVidrio10").price
        .Cells(6, 3).value = database.getMaterial("precioVidrio2").price
        .Cells(7, 3).value = database.getMaterial("precioM525").price
    End With
    ExcelApp.ActiveWorkbook.Close (1)
    Call database.closeConectionDB
End Sub
Sub moveFile(sourceFile, targetFile)
    Set FSO = CreateObject("Scripting.FileSystemObject")
    If FSO.FileExists(targetFile) Then
        Debug.Print "El archivo " & directoryPath & " ya existe."
    ElseIf FSO.FileExists(sourceFile) Then
        FSO.moveFile sourceFile, targetFile
    End If
    Set FSO = Nothing
End Sub

Sub copyFile(sourceFile, targetFile)
    Set FSO = CreateObject("Scripting.FileSystemObject")
    If FSO.FileExists(sourceFile) Then
        FSO.copyFile sourceFile, targetFile, True
        Debug.Print "El archivo " & directoryPath & " ya existe."
    End If
    Set FSO = Nothing
End Sub

Sub closeExcel1()
    On Error Resume Next
    While Err.Number = 0
        Set objOffice = GetObject(, "Excel.Application")
        objOffice.DisplayAlerts = False
        For Each objWindow In objOffice.Windows
            objWindow.Activate
            Set WBook = objOffice.ActiveWorkbook
            WBook.Saved = True
            WBook.Close
        Next
        objOffice.DisplayAlerts = True
        objOffice.Quit
        Set objOffice = Nothing
        WScript.Sleep 2000
    Wend
    
    MsgBox "Done"
End Sub

Sub closeExcel()
    Shell ("cmd.exe /c taskkill /im ""EXCEL.exe"" ")
End Sub

Sub closeInventor()

    Dim invApp As Inventor.Application
    On Error Resume Next
    Set invApp = GetObject(, "Inventor.Application")
    Dim IsInventorRunning As Boolean
    IsInventorRunning = (Err.Number = 0)
    If IsInventorRunning Then
        invApp.ActiveDocument.Close True
        invApp.Quit
        Set invApp = Nothing
    End If
    Set invApp = Nothing
    Err.Clear
End Sub


Sub closeWord()
    Dim WordApp As Word.Application
    On Error Resume Next
    Set WordApp = GetObject(, "Word.Application")
    Dim isWordAppRunning As Boolean
    isWordAppRunning = (Err.Number = 0)
    
    If isWordAppRunning Then
        WordApp.Quit
    End If
    Set WordApp = Nothing
    Err.Clear
End Sub
