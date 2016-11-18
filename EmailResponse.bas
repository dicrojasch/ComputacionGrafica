Attribute VB_Name = "EmailResponse"
Public Sub InMail(mail As Outlook.MailItem)
    
    Dim tiempo As New CalculateTime
    tiempo.StartTimer
        
    Dim objetoJson As Object
    Dim cliente As New Client
    Dim quot As New Quote
    Dim producto As New Product
    Dim formaleta As Formaletas
    Dim invernadero As Invernaderos
                
    Set objetoJson = parseJSON(mail.body)
    Set quot.cliente = New Client
    Call quot.cliente.JSONtoClient(objetoJson)
    
    quot.benefit = 0.2
    
    If Not objetoJson Is Nothing Then
        If "formaleta" = objetoJson.item("formulario") Then
        
            Set formaleta = New Formaletas
            Call formaleta.JSONtoFormaleta(objetoJson)
            Set quot.producto = New Product
            Call quot.producto.setFormaleta(formaleta)
            quot.producto.price = 2000000
                    ' To Do: Calculate price in product with materiales
            Call addExcel.pasarAExcelFormaleta(formaleta)
            
        ElseIf "invernadero" = objetoJson.item("formulario") Then
            
            Set invernadero = New Invernaderos
            Call invernadero.JSONtoInvernaderos(objetoJson)
            Set quot.producto = New Product
            Call quot.producto.setInvernadero(invernadero)
            quot.producto.price = 5000000
                    ' To Do: Calculate price in product with materiales
            Call addExcel.pasarExcelInvernadero(invernadero)
                    
        End If
        
        Dim database As New GraficaDB
        Call database.ConnectDB(DBServer, schema, user, password)
        
        Set quot.cliente = database.CreateClient(quot.cliente)
        If cliente.id = 0 Then
            Debug.Print "No se creo cliente"
        End If
        
        Set quot.producto = database.CreateProduct(quot.producto)
        If producto.id = 0 Then
            Debug.Print "No se creo producto"
        End If
        
        ' State 1, The client and product exists in database
        quot.state = 1
        
        quot.time_response = tiempo.EndTimer
        If Not database.CreateQuote(quot) Then
            Debug.Print "No se creo cotizacion"
        End If
        
        
        Call Mail_Recieve(quot)
        ' State 2, The receive answer has been sent to the client
        quot.state = 2
        quot.time_response = tiempo.EndTimer
        If Not database.CreateQuote(quot) Then
            Debug.Print "No se creo cotizacion"
        End If
        
'        OpenInventorFile (path & pathExample)
        ' ToDo Inventor operations
        Call wordCotizacion(quot)
                
        ' State 3, the files has been created
        quot.state = 3
        quot.time_response = tiempo.EndTimer
        If Not database.UpdateQuote(quot) Then
            Debug.Print "No se creo cotizacion"
        End If
        
        
        Call Mail_Quote(quot)
        
        ' State 4, the answer to the client has been sent
        quot.state = 4
        quot.time_response = tiempo.EndTimer
        If Not database.UpdateQuote(quot) Then
            Debug.Print "No se creo cotizacion"
        End If
        
        Dim newDirectory As String
        newDirectory = path & "Cotizaciones\" & Year(Date) & "\" & getMonth & "\" & quot.cliente.firstName & "_" & quot.cliente.lastname & "_P" & quot.producto.id & "\"
        createDirectory (newDirectory)
        Call moveFile(path & "modelo2d.pdf", newDirectory & "modelo2d.pdf")
        Call copyFile(path & "Plantilla de datos.xlsx", newDirectory & "Plantilla de datos.xlsx")
        
        
        ' State 5, The files has been moved to the appropriate folders ans the operation has finished correctly
        quot.state = 5
        quot.time_response = tiempo.EndTimer
        If Not database.UpdateQuote(quot) Then
            Debug.Print "No se creo cotizacion"
        End If

        'TODO : database material, provider, purchase
        Call database.closeConectionDB
        
    End If
        
End Sub


' Funcion para reempalazar los campos que trae la plantilla del correo por los campos obtenidos del correo de la solicitud'
Function ChangeBody(body As String, quot As Quote) As String
    body = replace(body, "<<clientname>>", quot.cliente.firstName & " " & quot.cliente.lastname)
    body = replace(body, "<<producto>>", quot.producto.getName)
    body = replace(body, "<<parameters>>", quot.producto.mailDescription)
    body = replace(body, "<<date>>", getDate())
    body = replace(body, "<<price>>", str(quot.getPrice))
    ChangeBody = body
End Function

Sub Mail_Recieve(quot As Quote)
    Dim OutApp As Object
    Dim OutMail As Object
    Set OutApp = CreateObject("Outlook.Application")
    Set OutMail = OutApp.CreateItemFromTemplate(path & "Comprobante.oft")
    On Error Resume Next
    With OutMail
        .To = quot.cliente.email
        .Subject = "Cotizacion " & quot.producto.getName
        .body = ChangeBody(.body, quot)
        .Attachments.Add ActiveWorkbook.FullName
        .Send
    End With
    
    On Error GoTo 0
    Set OutMail = Nothing
    Set OutApp = Nothing
End Sub



Sub Mail_Quote(quot As Quote)
    Dim OutApp As Object
    Dim OutMail As Object
    Set OutApp = CreateObject("Outlook.Application")   'Crea un objeto Outlook'
    Set OutMail = OutApp.CreateItemFromTemplate(path & pathPlantilla)   'Mediante el objeto outlook crea un cuerpo de mensaje con el nombre de la plantilla que hayamos creado'
    On Error Resume Next
    With OutMail
        .To = quot.cliente.email                     'Especifica el correo al que se envia la respuesta'
        .Subject = "Cotizacion " & quot.producto.getName         'Especifica el asunto del correo'
        .body = ChangeBody(.body, quot)                         'Cambia el cuerpo de la plantilla donde se completa con los datos de la estructura quot'
        .Attachments.Add ActiveWorkbook.FullName        'Adjunta el correo la informacion especificada anteriormente'
        .Attachments.Add (path & "cotizacion" & quot.producto.id & ".pdf")          'Adjunta un archivo'
        .Attachments.Add (path & "Parametros_Invernaderos.xlsm")          'Adjunta un archivo'
        .Attachments.Add (path & modelo3D)          'Adjunta un archivo'
        .Send                                                                           'envia el correo'
    End With
    
    On Error GoTo 0
    Set OutMail = Nothing
    Set OutApp = Nothing
End Sub

