Attribute VB_Name = "TestModule"
Sub InMail1()

    Dim tiempo As New CalculateTime
    tiempo.StartTimer
    Dim body As String
    'body = "{""formulario"":""formaleta"",""medidas"":{""unidades"":""mm"",""altura"":10,""diametroInterno"":10,""alturaRanura"":10},""opciones"":{""CP_0"":{""activado"":true,""texto"":""W16X26""},""RV_0_90"":{""activado"":true}},""datosUsuario"":{""nombre"":""diego"",""apellidos"":""rojas"",""email"":""icquirogac@unal.edu.co""}}"
    body = "{""formulario"":""invernadero"",""medidas"":{""unidades"":""metros"",""tipo"":""piramide"",""ancho"":10,""largo"":10,""alto"":10},""datosUsuario"":{""nombre"":""DIego"",""apellidos"":""Rojas"",""email"":""asc@gmail.com""}}"
    Dim objetoJson As Object
    Dim cliente As New Client
    Dim quot As New Quote
    Dim producto As New Product
    Dim formaleta As Formaletas
    Dim invernadero As Invernaderos
                
    Set objetoJson = parseJSON(body)
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
        
        quot.producto.addMaterials
        
        
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
        
        If quot.producto.is_Invernadero Then
            Call ExecInvernaderos
        ElseIf quot.producto.is_Formaleta Then
            'Call ExecFormaletas
        End If
    
        Call wordCotizacion(quot)
                
        ' State 3, the files has been created
        quot.state = 3
        quot.time_response = tiempo.EndTimer
        If Not database.UpdateQuote(quot) Then
            Debug.Print "No se creo cotizacion"
        End If
        
        
        'Call Mail_Quote(quot)
        
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
        Call moveFile(path & "cotizacion" & quot.producto.id & ".pdf", newDirectory & "cotizacion" & quot.producto.id & ".pdf")
        
        
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

Public Sub test()
'    Dim test As GraficaDB
'    Set test = New GraficaDB
'
'    Dim cliente As Client
'    cliente.firstName = "diego"
'    cliente.lastName = "rojas"
'    cliente.email = "dicrojasch@unal.edu.co"
'    Dim getcli As Client
'
'    Call test.ConnectDB("127.0.0.1", "grafica", "root", "dcrojas.3124")
'
'    getcli = test.CreateClient(cliente)
'    MsgBox getcli.id & ", " & getcli.firstName
'
'    Call test.closeConectionDB
    
    
    
    

''--------------------------------------------------------------------


'
'    Dim formaleta As New Formaletas
'    Call formaleta.InitFormaletas(10, 10, 10)
'    formaleta.aFPlate0 = "012354678910"
'    formaleta.cPlate0 = "012354678910"
'    formaleta.cPlate90 = "012354678910"
'    formaleta.cPlate180 = "012354678910"
'    formaleta.cPlate270 = "012354678910"
'    formaleta.aFPlate0 = "012354678910"
'    formaleta.aFPlate45 = "012354678910"
'    formaleta.aFPlate90 = "012354678910"
'    formaleta.aFPlate135 = "012354678910"
'    formaleta.aFPlate180 = "012354678910"
'    formaleta.aFPlate225 = "012354678910"
'    formaleta.aFPlate270 = "012354678910"
'    formaleta.aFPlate315 = "012354678910"
'    formaleta.rVar0_90 = True
'    formaleta.rVar90_180 = True
'    formaleta.rVar180_270 = True
'    formaleta.rVar270_0 = True
'
'    Dim prod As New Product
'    Call prod.setFormaleta(formaleta)
'    prod.price = 100000
'

'    Dim Invern As New Invernaderos
'    Call Invern.setAreaLado(10, 10)
'    Dim prod1 As New Product
'    Call prod1.setInvernadero(Invern)
'    prod1.price = 320000
'
'
'
'    Dim test As GraficaDB
'    Set test = New GraficaDB
'    Call test.ConnectDB("127.0.0.1", "grafica", "root", "dcrojas.3124")
'    Set prod1 = test.CreateProduct(prod1)
'    MsgBox prod1.id
'    'MsgBox test.CreateProduct(prod)
'    Call test.closeConectionDB




    
'    Dim Invern As New Invernaderos
'    Call Invern.setAreaLado(10, 10)
'
'    Dim prod1 As New Product
'    Call prod1.setInvernadero(Invern)
'    prod1.price = 320000
'   ' MsgBox Date$
'   Dim quot As Quote
'
'   quot.price = prod1.price * 1.2
'    Dim test As GraficaDB
'    Set test = New GraficaDB
'    Call test.ConnectDB("127.0.0.1", "grafica", "root", "dcrojas.3124")
'    quot.cliente = test.getClient("diego", "rojas")
'    Set quot.prod = test.CreateProduct(prod1)
'    MsgBox test.CreateQuote(quot)
'    Call test.closeConectionDB
'
'
'    Dim mat As New Material
'
'
'    mat.name = "Cuero"
'    mat.description = "delgado"
'    mat.price = 100000
'    mat.quantity = 100
'
'    Dim test As GraficaDB
'    Set test = New GraficaDB
'    Call test.ConnectDB("127.0.0.1", "grafica", "root", "dcrojas.3124")
'    MsgBox test.CreateMaterial(mat)
'    Call test.closeConectionDB
    
    
    
'        Dim material1 As New Material
'    Dim material2 As New Material
'    material1.name = "material1"
'    material2.name = "material2"
'
'    Dim materials As Dictionary
'    Set materials = New Dictionary
'    Call materials.Add(material1, 1)
'
'    MsgBox materials(material2)


'------------------------------- Ejemplo completo
'    Dim calTime As New CalculateTime
'    Dim calTime2 As New CalculateTime
'    calTime.StartTimer
'    calTime2.StartTimer
'
'
'    Dim cliente As New Client
'    Dim getcli As New Client
'    cliente.firstName = "diego2"
'    cliente.lastname = "rojas"
'    cliente.email = "dicrojasch@unal.edu.co"
'
'    Dim formaleta As New Formaletas
'    Call formaleta.InitFormaletas(10, 10, 10)
'    formaleta.aFPlate0 = "012354678910"
'    formaleta.cPlate0 = "012354678910"
'    formaleta.cPlate90 = "012354678910"
'    formaleta.cPlate180 = "012354678910"
'    formaleta.cPlate270 = "012354678910"
'    formaleta.aFPlate0 = "012354678910"
'    formaleta.aFPlate45 = "012354678910"
'    formaleta.aFPlate90 = "012354678910"
'    formaleta.aFPlate135 = "012354678910"
'    formaleta.aFPlate180 = "012354678910"
'    formaleta.aFPlate225 = "012354678910"
'    formaleta.aFPlate270 = "012354678910"
'    formaleta.aFPlate315 = "012354678910"
'    formaleta.rVar0_90 = True
'    formaleta.rVar90_180 = True
'    formaleta.rVar180_270 = True
'    formaleta.rVar270_0 = True
'    Dim prodf As New Product
'    Call prodf.setFormaleta(formaleta)
'    prodf.price = 100000
'
'    Dim Invern As New Invernaderos
'    Call Invern.setAreaLado(10, 10)
'    Dim prodi As New Product
'    Call prodi.setInvernadero(Invern)
'    prodi.price = 320000
'
'    Dim getProd As New Product
'
'    Dim cemento As New Material
'    Dim varilla As New Material
'    Dim lamina As New Material
'    cemento.name = "Cemento"
'    cemento.description = "marca x"
'    cemento.quantity = 100
'    cemento.quantity_purch = 0
'    varilla.name = "varilla"
'    varilla.description = "marca y"
'    varilla.quantity = 100
'    varilla.quantity_purch = 0
'    lamina.name = "lamina"
'    lamina.description = "marca l"
'    lamina.quantity = 100
'    lamina.quantity_purch = 0
'    Dim check As Boolean
'    check = prodf.addMaterial(cemento)
'    check = prodf.addMaterial(varilla)
'    check = prodi.addMaterial(varilla)
'    check = prodi.addMaterial(lamina)
'
'    Dim prov As New Provider
'    prov.name = "Industrias Colombia S.A."
'    prov.email = "ind@ind.com"
'
'    Dim purch As New Purchases
'    purch.provider_name = prov.name
'    purch.description = "test"
'    check = purch.addMaterial(cemento)
'    check = purch.addMaterial(varilla)
'    check = purch.addMaterial(lamina)
'    Dim isCreated As Boolean
'
'    Dim quotf As New Quote
'    Dim quoti As New Quote
'
'
'
'    Dim test As GraficaDB
'    Set test = New GraficaDB
'    Call test.ConnectDB("127.0.0.1", "grafica", "root", "dcrojas.3124")
'    Set cliente = test.CreateClient(cliente)
'
'    isCreated = test.CreateMaterial(cemento)
'    isCreated = test.CreateMaterial(varilla)
'    isCreated = test.CreateMaterial(lamina)
'
'    isCreated = test.CreateProvider(prov)
'
'    isCreated = test.CreatePurchase(purch)
'
'    Set prodf = test.CreateProduct(prodf)
'    Set prodi = test.CreateProduct(prodi)
'
'    Set quotf.cliente = cliente
'    Set quotf.producto = prodf
'    quotf.benefit = 0.2
'
'    Set quoti.cliente = cliente
'    Set quoti.producto = prodi
'    quoti.benefit = 0.3
'    Debug.Print calTime2.EndTimer
'    Pause (2)
'
'
'    quotf.time_response = calTime.EndTimer
'    quoti.time_response = calTime.EndTimer
'    isCreated = test.CreateQuote(quotf)
'    isCreated = test.CreateQuote(quoti)
'
'    Call test.closeConectionDB
'
'    MsgBox "finalizo"
'    Call test.ConnectDB("127.0.0.1", "grafica", "root", "dcrojas.3124")
'    getcli = test.CreateClient(cliente)
'    Call test.closeConectionDB
'
'    Dim ab As New Dictionary
'    Call ab.Add(1, 2)
'    Call ab.Add(2, 3)
'    Call ab.Add(3, 4)
'    Call ab.Add(4, 5)
'    For Each a In ab.Items
'        MsgBox a
'    Next a

'    Dim cal As Date
'    cal = Now
'    Pause (20)
'
'    cal = Now - cal
'    Debug.Print cal
'


End Sub

Sub testClient()
    
End Sub




