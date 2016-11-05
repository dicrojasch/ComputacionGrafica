Attribute VB_Name = "EmailResponse"

Sub InMail(mail As Outlook.MailItem) '
    Dim tiempo As New CalculateTime
    tiempo.StartTimer
    
    Dim objetoJson As Object
    'Dim cliente As New Client
    Dim quot As New Quote
    Dim producto As New Product
    Dim formaleta As Formaletas
    Dim invernadero As Invernaderos
                
    Set objetoJson = parseJSON(mail.body)
    cliente.JSONtoClient (objetoJson)
    quot.cliente.JSONtoClient (objetoJson)
    
    quot.benefit = 0.2
    
    If Not objetoJson Is Nothing Then
        If "formaleta" = objetoJson.Item("formulario") Then
        
            Set formaleta = New Formaletas
            Call formaleta.JSONtoFormaleta(objetoJson)
            quot.producto.setFormaleta (formaleta)
            Call ExcelFormaleta.pasarAExcelFormaleta(objetoJson)
            
        ElseIf "invernadero" = objetoJson.Item("formulario") Then
            
            Set invernadero = New Invernaderos
            invernadero.JSONtoInvernaderos (objetoJson)
            quot.producto.setInvernadero (invernadero)
            ' TODO : excel
            
            
        End If
        Call Mail_Quote(quot)
        
        OpenInventorFile ("C:\Users\diego\Desktop\Example Inventor\12\toExporti.iam")
            
        Dim database As New GraficaDB
        Call database.ConnectDB("127.0.0.1", "grafica", "root", "dcrojas.3124")
        database.CreateClient (quot.cliente)
        database.CreateProduct (quot.producto)
        quot.time_response = tiempo.EndTimer
        database.CreateQuote (quot)
        'TODO : database material, provider, purchase
        Call database.closeConectionDB
        
    End If
    
    
    
    
    
End Sub


' Funcion para reempalazar los campos que trae la plantilla del correo por los campos obtenidos del correo de la solicitud'
Function ChangeBody(body As String, quot As Quote) As String
    body = Replace(body, "<<clientname>>", quot.cliente.firstName)
    'body = Replace(body, "<<producto>>", quot.product)
    'body = Replace(body, "<<parameters>>", quot.parameters)
    body = Replace(body, "<<date>>", getDate())
    'body = Replace(body, "<<price>>", Str(quot.price))
    ChangeBody = body
End Function



' Crea y enviar un correo con archivos adjuntos'
Sub Mail_Quote(quot As Quote)
    Dim OutApp As Object
    Dim OutMail As Object
    Set OutApp = CreateObject("Outlook.Application")   'Crea un objeto Outlook'
    Set OutMail = OutApp.CreateItemFromTemplate(path & "cotizacion.oft")   'Mediante el objeto outlook crea un cuerpo de mensaje con el nombre de la plantilla que hayamos creado'
    OpenInventorFile ("C:\Users\diego\Desktop\Example Inventor\12\toExporti.iam")
    On Error Resume Next
    With OutMail
        .To = quot.cliente.email                                                      'Especifica el correo al que se envia la respuesta'
        .Subject = "Cotizacion " '& quot.product         'Especifica el asunto del correo'
        .body = ChangeBody(.body, quot)                         'Cambia el cuerpo de la plantilla donde se completa con los datos de la estructura quot'
        .Attachments.Add ActiveWorkbook.FullName        'Adjunta el correo la informacion especificada anteriormente'
        .Attachments.Add (path & "Plantilla de datos.xlsx")          'Adjunta un archivo'
        .Attachments.Add (path & "modelo2d.xlsx")          'Adjunta un archivo'
        .Send                                                                           'envia el correo'
        MsgBox "funciona: " & .Sent
    End With
    
    On Error GoTo 0
    Set OutMail = Nothing
    Set OutApp = Nothing
End Sub




