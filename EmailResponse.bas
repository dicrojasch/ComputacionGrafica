Attribute VB_Name = "EmailResponse"
Public Sub InMail(mail As Outlook.MailItem)
    Call closeInventor
    Call moveInvernaderoFiles(path & "Dropbox\Missing Files")
    Call moveFormaletaFiles(path & "Dropbox\Missing Files")
 
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
            Call addExcel.pasarAExcelFormaleta(formaleta, quot.producto.id)
            
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
        
   
        ' State 2, The receive answer has been sent to the client
        quot.state = 2
        quot.time_response = tiempo.EndTimer
        If Not database.CreateQuote(quot) Then
            Debug.Print "No se creo cotizacion"
        End If
'        MsgBox (quot.producto.getName)
'        MsgBox (quot.producto.is_Formaleta)
'        MsgBox (quot.producto.is_Invernadero)
        If quot.producto.is_Invernadero Then
            Call ExecInvernaderos
        ElseIf quot.producto.is_Formaleta Then
            Call ExecFormaletas
        End If
                
        ' State 3, the files has been created
        quot.producto.addMaterials
        quot.state = 3
        quot.time_response = tiempo.EndTimer
        If Not database.UpdateQuote(quot) Then
            Debug.Print "No se creo cotizacion"
        End If
        
        MsgBox ("e inventor qpaso?")
        Call Mail_Quote(quot)
        
        ' State 4, the answer to the client has been sent
        quot.state = 4
        quot.time_response = tiempo.EndTimer
        If Not database.UpdateQuote(quot) Then
            Debug.Print "No se creo cotizacion"
        End If
        
        Dim newDirectory As String
        newDirectory = path & "Dropbox\Cotizaciones\" & Year(Date) & "\" & getMonth & "\" & quot.cliente.firstName & "_" & quot.cliente.lastname & "_P" & quot.producto.id & "\"
        createDirectory (newDirectory)
  
        If quot.producto.is_Formaleta Then
            Call moveFormaletaFiles(newDirectory)
        ElseIf quot.producto.is_Invernadero Then
            Call moveInvernaderoFiles(newDirectory)
        End If
        'Call moveFile(path & "cotizacion" & quot.producto.id & ".pdf", newDirectory & "cotizacion" & quot.producto.id & ".pdf")
        
        
        ' State 5, The files has been moved to the appropriate folders ans the operation has finished correctly
        quot.state = 5
        quot.time_response = tiempo.EndTimer
        If Not database.UpdateQuote(quot) Then
            Debug.Print "No se creo cotizacion"
        End If

        'TODO : database material, provider, purchase
        Call database.closeConectionDB
        
    End If
    Call closeInventor
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
        If quot.producto.is_Invernadero Then
            .Attachments.Add (pathInvernaderos & "Estructura_BT4D.jpg")          'Adjunta un archivo'
            .Attachments.Add (pathInvernaderos & "Estructura_BT7D.jpg")
            .Attachments.Add (pathInvernaderos & "Estructura_RC.jpg")
            .Attachments.Add (pathInvernaderos & "Estructura_RT.jpg")
            .Attachments.Add (pathInvernaderos & "Cotizacion_BT4.pdf")
            .Attachments.Add (pathInvernaderos & "Cotizacion_BT7.pdf")
            .Attachments.Add (pathInvernaderos & "Cotizacion_RC.pdf")
            .Attachments.Add (pathInvernaderos & "Cotizacion_RT.pdf")
        ElseIf quot.producto.is_Formaleta Then
            .Attachments.Add (pathInvernaderos & "Cotizacion_Formaleta.pdf")
            .Attachments.Add (pathInvernaderos & "Plano 1.pdf")
            .Attachments.Add (pathInvernaderos & "Plano 2.pdf")
            .Attachments.Add (pathInvernaderos & "FORMALETA BASE.jpg")
        End If
        .Send                                                                           'envia el correo'
    End With
    
    On Error GoTo 0
    Set OutMail = Nothing
    Set OutApp = Nothing
End Sub


Sub moveFormaletaFiles(newDirectory As String)
    Call moveFile(pathInvernaderos & "Cotizacion_Formaleta.pdf", newDirectory & "Cotizacion_Formaleta.pdf")
    Call moveFile(pathInvernaderos & "Lista de Materiales.xlsx", newDirectory & "Lista de Materiales.xlsx")
    Call moveFile(pathInvernaderos & "Plano 1.pdf", newDirectory & "Plano 1.pdf")
    Call moveFile(pathInvernaderos & "Plano 2.pdf", newDirectory & "Plano 2.pdf")
    Call moveFile(pathInvernaderos & "FORMALETA BASE.jpg", newDirectory & "FORMALETA BASE.jpg")
End Sub


Sub moveInvernaderoFiles(newDirectory As String)
            Call moveFile(pathInvernaderos & "Estructura_BT4D.jpg", newDirectory & "Estructura_BT4D.jpg")
            Call moveFile(pathInvernaderos & "Estructura_BT7D.jpg", newDirectory & "Estructura_BT7D.jpg")
            Call moveFile(pathInvernaderos & "Estructura_RC.jpg", newDirectory & "Estructura_RC.jpg")
            Call moveFile(pathInvernaderos & "Estructura_RT.jpg", newDirectory & "Estructura_RT.jpg")
            Call moveFile(pathInvernaderos & "Cotizacion_BT4.pdf", newDirectory & "Cotizacion_BT4.pdf")
            Call moveFile(pathInvernaderos & "Cotizacion_BT7.pdf", newDirectory & "Cotizacion_BT7.pdf")
            Call moveFile(pathInvernaderos & "Cotizacion_RC.pdf", newDirectory & "Cotizacion_RC.pdf")
            Call moveFile(pathInvernaderos & "Cotizacion_RT.pdf", newDirectory & "Cotizacion_RT.pdf")
            Call moveFile(pathInvernaderos & "BOM_BT4.xlsx", newDirectory & "BOM_BT4.xlsx")
            Call moveFile(pathInvernaderos & "BOM_BT7.xlsx", newDirectory & "BOM_BT7.xlsx")
            Call moveFile(pathInvernaderos & "BOM_RC.xlsx", newDirectory & "BOM_RC.xlsx")
            Call moveFile(pathInvernaderos & "BOM_RT.xlsx", newDirectory & "BOM_RT.xlsx")
            ' ToDo : Move qote files
End Sub
