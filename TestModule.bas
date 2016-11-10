Attribute VB_Name = "TestModule"
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
    Dim calTime As New CalculateTime
    Dim calTime2 As New CalculateTime
    calTime.StartTimer
    calTime2.StartTimer


    Dim cliente As New Client
    Dim getcli As New Client
    cliente.firstName = "diego2"
    cliente.lastname = "rojas"
    cliente.email = "dicrojasch@unal.edu.co"

    Dim formaleta As New Formaletas
    Call formaleta.InitFormaletas(10, 10, 10)
    formaleta.aFPlate0 = "012354678910"
    formaleta.cPlate0 = "012354678910"
    formaleta.cPlate90 = "012354678910"
    formaleta.cPlate180 = "012354678910"
    formaleta.cPlate270 = "012354678910"
    formaleta.aFPlate0 = "012354678910"
    formaleta.aFPlate45 = "012354678910"
    formaleta.aFPlate90 = "012354678910"
    formaleta.aFPlate135 = "012354678910"
    formaleta.aFPlate180 = "012354678910"
    formaleta.aFPlate225 = "012354678910"
    formaleta.aFPlate270 = "012354678910"
    formaleta.aFPlate315 = "012354678910"
    formaleta.rVar0_90 = True
    formaleta.rVar90_180 = True
    formaleta.rVar180_270 = True
    formaleta.rVar270_0 = True
    Dim prodf As New Product
    Call prodf.setFormaleta(formaleta)
    prodf.price = 100000

    Dim Invern As New Invernaderos
    Call Invern.setAreaLado(10, 10)
    Dim prodi As New Product
    Call prodi.setInvernadero(Invern)
    prodi.price = 320000

    Dim getProd As New Product

    Dim cemento As New Material
    Dim varilla As New Material
    Dim lamina As New Material
    cemento.name = "Cemento"
    cemento.description = "marca x"
    cemento.quantity = 100
    cemento.quantity_purch = 0
    varilla.name = "varilla"
    varilla.description = "marca y"
    varilla.quantity = 100
    varilla.quantity_purch = 0
    lamina.name = "lamina"
    lamina.description = "marca l"
    lamina.quantity = 100
    lamina.quantity_purch = 0
    Dim check As Boolean
    check = prodf.addMaterial(cemento)
    check = prodf.addMaterial(varilla)
    check = prodi.addMaterial(varilla)
    check = prodi.addMaterial(lamina)

    Dim prov As New Provider
    prov.name = "Industrias Colombia S.A."
    prov.email = "ind@ind.com"

    Dim purch As New Purchases
    purch.provider_name = prov.name
    purch.description = "test"
    check = purch.addMaterial(cemento)
    check = purch.addMaterial(varilla)
    check = purch.addMaterial(lamina)
    Dim isCreated As Boolean

    Dim quotf As New Quote
    Dim quoti As New Quote



    Dim test As GraficaDB
    Set test = New GraficaDB
    Call test.ConnectDB("127.0.0.1", "grafica", "root", "dcrojas.3124")
    Set cliente = test.CreateClient(cliente)

    isCreated = test.CreateMaterial(cemento)
    isCreated = test.CreateMaterial(varilla)
    isCreated = test.CreateMaterial(lamina)

    isCreated = test.CreateProvider(prov)

    isCreated = test.CreatePurchase(purch)

    Set prodf = test.CreateProduct(prodf)
    Set prodi = test.CreateProduct(prodi)

    Set quotf.cliente = cliente
    Set quotf.producto = prodf
    quotf.benefit = 0.2

    Set quoti.cliente = cliente
    Set quoti.producto = prodi
    quoti.benefit = 0.3
    Debug.Print calTime2.EndTimer
    Pause (2)


    quotf.time_response = calTime.EndTimer
    quoti.time_response = calTime.EndTimer
    isCreated = test.CreateQuote(quotf)
    isCreated = test.CreateQuote(quoti)

    Call test.closeConectionDB

    MsgBox "finalizo"
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
    
    Call moveFile(path & "modelo2d.pdf", path & "prueba/modelo2d.pdf")
    
    
End Sub
