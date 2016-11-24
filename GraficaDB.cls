VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "GraficaDB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private GraficaConnect As ADODB.Connection


Public Function checkConnect() As Boolean
    If GraficaConnect Is Nothing Then
        checkConnect = False
    Else
        checkConnect = True
    End If

End Function
    
Public Function ConnectDB(server As String, database As String, user As String, password As String)
    Set GraficaConnect = New ADODB.Connection
    GraficaConnect.Open "Driver={MySQL ODBC 5.3 ANSI Driver};Server=" & server & ";Database=" & database & ";User=" & user & ";Password=" & password & ";Option=3;"
End Function


Public Sub closeConectionDB()
    If checkConnect Then
        GraficaConnect.Close
        Set GraficaConnect = Nothing
    Else
        Debug.Print "Actualmente no hay una conexion con base de datos."
    End If
End Sub

Public Function getClient(firstName As String, lastname As String, email As String) As Client
    Set getClient = New Client
    If checkConnect Then
        Dim query As String
        Dim answer As Variant
        Set records = New ADODB.Recordset
        query = "Select * from cliente where name = '" & firstName & "'And lastname = '" & lastname & "' And email = '" & email & "';"
        records.Open query, GraficaConnect, adOpenDynamic
        If Not records.EOF Then
            answer = records.GetRows()
            getClient.id = answer(0, 0)
            getClient.firstName = answer(1, 0)
            getClient.lastname = answer(2, 0)
            getClient.email = answer(3, 0)
        End If
        records.Close
        Set records = Nothing
    Else
        Debug.Print "No hay conexion con la base de datos."
    End If
End Function


Public Function CreateClient(Client1 As Client) As Client
    If checkConnect Then
        Dim tempClient As New Client
        Set tempClient = getClient(Client1.firstName, Client1.lastname, Client1.email)
        Set CreateClient = tempClient
        If tempClient.firstName = "" And tempClient.lastname = "" And tempClient.id = 0 Then
            Dim query As String
            Set records = New ADODB.Recordset
            query = "select * from cliente;"
            records.Open query, GraficaConnect, adOpenDynamic, adLockOptimistic
            records.AddNew
            records!name = Client1.firstName
            records!lastname = Client1.lastname
            records!email = Client1.email
            records.Update
            records.Close
            Set records = Nothing
            Set CreateClient = getClient(Client1.firstName, Client1.lastname, Client1.email)
        End If
    Else
        Debug.Print "No hay conexion con la base de datos."
    End If
End Function


Private Function getLastProductId() As Integer
    If checkConnect Then
        Dim query As String
        Set records = New ADODB.Recordset
        query = "SELECT MAX(id_product) FROM product;"
        records.Open query, GraficaConnect, adOpenDynamic
        If Not records.EOF Then
            getLastProductId = records.Fields(0).value
        End If
        records.Close
        Set records = Nothing
    Else
        Debug.Print "No hay conexion con la base de datos."
    End If

End Function

Public Function CreateProduct(product1 As Product) As Product
    If checkConnect Then
        Dim query As String
        Set records = New ADODB.Recordset
        query = "select * from product;"
        records.Open query, GraficaConnect, adOpenDynamic, adLockOptimistic
        records.AddNew
        records!name = product1.getName
        records!description = product1.getDescription
        records!price = product1.price
        records.Update
        records.Close
        Set records = Nothing
        product1.id = getLastProductId()
        
        If Not CreateMaterialProduct(product1) Then
            Debug.Print "Error al crear producto: el producto no tiene materiales"
        End If
        Set CreateProduct = product1
    Else
        Debug.Print "No hay conexion con la base de datos."
    End If
End Function


Public Function checkProduct(product1 As Product) As Boolean
    If checkConnect Then
        Dim query As String
        Dim answer As Variant
        Set records = New ADODB.Recordset
        query = "Select * from product where id_product = '" & product1.id & "';"
        records.Open query, GraficaConnect, adOpenDynamic
        If Not records.EOF Then
            answer = records.GetRows()
            Dim requirements As Boolean
            requirements = True
            requirements = requirements And product1.id = answer(0, 0)
            requirements = requirements And product1.getName = answer(1, 0)
            requirements = requirements And product1.getDescription = answer(2, 0)
            requirements = requirements And product1.price = answer(3, 0)
            checkProduct = requirements
        Else
            checkProduct = False
        End If
        records.Close
        Set records = Nothing
    End If
End Function

Private Function CreateMaterialProduct(product1 As Product) As Boolean
    If checkConnect Then
        If checkProduct(product1) Then
            Dim query As String
            query = "select * from material_product;"
            Dim materials As New Collection
            Set materials = product1.getAllMaterials
            If materials.Count > 0 Then
                For Each mat In materials
                    If checkMaterial(mat.name) Then
                        Set records = New ADODB.Recordset
                        records.Open query, GraficaConnect, adOpenDynamic, adLockOptimistic
                        records.AddNew
                        records!product_id = product1.id
                        records!material_name = mat.name
                        records!quantity = mat.quotQuantity
                        records.Update
                        records.Close
                        Set records = Nothing
                    Else
                        Debug.Print "Error al asociar materiales con productos: El material '" & mat.name & "' no existe en base de datos."
                    End If
                Next mat
                CreateMaterialProduct = True
            Else
                Debug.Print "Error al asociar materiales con productos: No hay materiales en el producto."
                CreateMaterialProduct = False
            End If
            CreateMaterialProduct = True
        Else
            Debug.Print "Error al asociar materiales con productos: Los datos no coinciden con la base de datos."
            CreateMaterialProduct = False
        End If
    Else
        Debug.Print "No hay conexion con la base de datos."
        CreateMaterialProduct = False
    End If
End Function


Public Function checkQuote(cliente_id As Integer, product_id As Integer) As Boolean
    If checkConnect Then
        Dim query As String
        Dim answer As Variant
        Set records = New ADODB.Recordset
        query = "Select * from quote where cliente_id = '" & cliente_id & "'And product_id = '" & product_id & "';"
        records.Open query, GraficaConnect, adLockOptimistic
        If Not records.EOF Then
            checkQuote = True
        Else
            checkQuote = False
        End If
        records.Close
        Set records = Nothing
    End If
End Function


Public Function UpdateQuote(quot As Quote) As Boolean
   If checkConnect Then
        If checkQuote(quot.cliente.id, quot.producto.id) Then
            Dim query As String
            Set records = New ADODB.Recordset
            query = "SELECT * FROM quote WHERE cliente_id=" & quot.cliente.id & " AND product_id=" & quot.producto.id & ";"
            records.Open query, GraficaConnect, adOpenStatic, adLockOptimistic
            records!Date = Date$
            records!price = quot.getPrice
            records!time_response = quot.time_response
            records!state = quot.state
            records.Update
            records.Close
            Set records = Nothing
            UpdateQuote = True
        Else
            Debug.Print "Error al actualizar cotizacion: El registro en cotizacion no existe."
            UpdateQuote = False
        End If
    Else
        Debug.Print "No hay conexion con la base de datos."
        UpdateQuote = False
    End If
End Function

Public Function CreateQuote(quot As Quote) As Boolean
    If checkConnect Then
        If Not checkQuote(quot.cliente.id, quot.producto.id) Then
            Dim query As String
            Set records = New ADODB.Recordset
            query = "select * from quote;"
            records.Open query, GraficaConnect, adOpenDynamic, adLockOptimistic
            records.AddNew
            records!cliente_id = quot.cliente.id
            records!product_id = quot.producto.id
            records!Date = Date$
            records!price = quot.getPrice
            records!time_response = quot.time_response
            records.Update
            records.Close
            Set records = Nothing
            CreateQuote = True
        Else
            Debug.Print "Error al crear cotizacion: El registro en cotizacion ya existe."
            CreateQuote = False
        End If
    Else
        Debug.Print "No hay conexion con la base de datos."
        CreateQuote = False
    End If
End Function


Public Function checkMaterial(name As String) As Boolean
    If checkConnect Then
        Dim query As String
        Set records = New ADODB.Recordset
        query = "Select * from material where name = '" & name & "';"
        records.Open query, GraficaConnect, adOpenDynamic
        If Not records.EOF Then
            checkMaterial = True
        Else
            checkMaterial = False
        End If
        records.Close
        Set records = Nothing
    End If
End Function

Public Function getMaterial(name As String) As Material
    If checkConnect Then
        Dim query As String
        Dim answer As Variant
        Set records = New ADODB.Recordset
        query = "Select * from material where name = '" & name & "';"
        records.Open query, GraficaConnect, adOpenDynamic
        If Not records.EOF Then
            Set getMaterial = New Material
            answer = records.GetRows()
            getMaterial.name = answer(0, 0)
            getMaterial.description = answer(1, 0)
            getMaterial.price = answer(2, 0)
            getMaterial.quantity = answer(3, 0)
            'ToDo : getMaterial.min = answer(4, 0)
        End If
        records.Close
        Set records = Nothing
    Else
        Debug.Print "No hay conexion con la base de datos."
    End If
End Function

Private Function addQuantityMat(mat As Material) As Material
    Set addQuantityMat = mat
    If checkConnect Then
        If mat.updateQuantity Then
            Dim query As String
            query = "Select * from material where name = '" & mat.name & "' ;"
            Set records = New ADODB.Recordset
            records.Open query, GraficaConnect, adOpenDynamic, adLockOptimistic
            If Not records.EOF Then
                records!quantity = mat.quantity
                records.Update
                Set addQuantityMat = mat
            End If
            records.Close
            Set records = Nothing
        End If
    Else
        Debug.Print "No hay conexion con la base de datos."
    End If

End Function

Public Function CreateMaterial(mat As Material) As Boolean
    If checkConnect Then
        If Not checkMaterial(mat.name) Then
            Dim query As String
            Set records = New ADODB.Recordset
            query = "select * from material;"
            records.Open query, GraficaConnect, adOpenDynamic, adLockOptimistic
            records.AddNew
            records!name = mat.name
            records!description = mat.description
            records!price = mat.price
            records!quantity = mat.quantity
            records.Update
            records.Close
            Set records = Nothing
            CreateMaterial = True
        Else
            Debug.Print "El material " & mat.name & " ya existe."
            CreateMaterial = False
        End If
    Else
        Debug.Print "No hay conexion con la base de datos."
        CreateMaterial = False
    End If
End Function



Public Function checkProvider(name As String) As Boolean
    If checkConnect Then
        Dim query As String
        Set records = New ADODB.Recordset
        query = "Select * from provider where name = '" & name & "';"
        records.Open query, GraficaConnect, adOpenDynamic
        If Not records.EOF Then
            checkProvider = True
        Else
            checkProvider = False
        End If
        records.Close
        Set records = Nothing
    End If
End Function

Public Function getProvider(name As String) As Provider
    If checkConnect Then
        Dim query As String
        Dim answer As Variant
        Set records = New ADODB.Recordset
        query = "Select * from material where name = '" & name & "';"
        records.Open query, GraficaConnect, adOpenDynamic
        If Not records.EOF Then
            answer = records.GetRows()
            getProvider.name = answer(0, 0)
            getProvider.email = answer(1, 0)
            getProvider.phone = answer(2, 0)
            getProvider.address = answer(3, 0)
        End If
        records.Close
        Set records = Nothing
    Else
        Debug.Print "No hay conexion con la base de datos."
    End If
End Function

Public Function CreateProvider(prov As Provider) As Boolean
    If checkConnect Then
        If Not checkProvider(prov.name) Then
            Dim query As String
            Set records = New ADODB.Recordset
            query = "select * from provider;"
            records.Open query, GraficaConnect, adOpenDynamic, adLockOptimistic
            records.AddNew
            records!name = prov.name
            records!email = prov.email
            records!phone = prov.phone
            records!address = prov.address
            records.Update
            records.Close
            Set records = Nothing
            CreateProvider = True
        Else
            Debug.Print "El Proveedor " & prov.name & " ya existe en base de datos."
            CreateProvider = False
        End If
    Else
        Debug.Print "No hay conexion con la base de datos."
        CreateProvider = False
    End If
End Function

Public Function CreatePurchase(purchase As Purchases) As Purchases
    If checkConnect Then
        If checkProvider(purchase.provider_name) And purchase.existsMaterials() Then
            Dim query As String
            Set records = New ADODB.Recordset
            query = "select * from purchases;"
            records.Open query, GraficaConnect, adOpenDynamic, adLockOptimistic
            records.AddNew
            records!provider_name = purchase.provider_name
            records!description = purchase.description
            records!Date = Date$
            records!total_price = purchase.getTotalPrice()
            records.Update
            records.Close
            Set records = Nothing
            purchase.id = getLastPurchaseId()
            If Not CreatePurchaseMaterial(purchase) Then
                Debug.Print "Error al crear compra: No se pudo asociar la compra con materiales"
            End If
            CreatePurchase = purchase
        Else
            Debug.Print "Error al crear compra: No se ha creado el proveedor o no hay materiales en la compra."
            CreatePurchase = Nothing
        End If
    Else
        Debug.Print "No hay conexion con la base de datos."
        CreatePurchase = Nothing
    End If
End Function

Public Function checkPurchase(purchase As Purchases) As Boolean
    If checkConnect Then
        Dim query As String
        Dim answer As Variant
        Set records = New ADODB.Recordset
        query = "Select * from purchases where id_purchases = '" & purchase.id & "';"
        records.Open query, GraficaConnect, adOpenDynamic
        If Not records.EOF Then
            answer = records.GetRows()
            Dim requirements As Boolean
            requirements = True
            requirements = requirements And purchase.id = answer(0, 0)
            requirements = requirements And purchase.provider_name = answer(1, 0)
            requirements = requirements And purchase.description = answer(2, 0)
            requirements = requirements And purchase.getTotalPrice = answer(4, 0)
            checkPurchase = requirements
        Else
            checkPurchase = False
        End If
        records.Close
        Set records = Nothing
    End If
End Function

Private Function getLastPurchaseId() As Integer
    If checkConnect Then
        Dim query As String
        Set records = New ADODB.Recordset
        query = "SELECT MAX(id_purchases) FROM purchases;"
        records.Open query, GraficaConnect, adOpenDynamic
        If Not records.EOF Then
            getLastPurchaseId = records.Fields(0).value
        End If
        records.Close
        Set records = Nothing
    Else
        Debug.Print "No hay conexion con la base de datos."
    End If

End Function

Private Function CreatePurchaseMaterial(purchase As Purchases) As Boolean
    If checkConnect Then
        If checkPurchase(purchase) Then
            Dim query As String
            query = "select * from purchases_material;"
            Dim materials As Collection
            Set materials = purchase.getAllMaterials
            Dim tmpMaterial As New Material
            If materials.Count > 0 Then
                For i = 1 To materials.Count
                    Set records = New ADODB.Recordset
                    records.Open query, GraficaConnect, adOpenDynamic, adLockOptimistic
                    Set tmpMaterial = materials.item(i)
                    records.AddNew
                    records!purchases_id = purchase.id
                    records!material_name = tmpMaterial.name
                    records!quantity = tmpMaterial.quantity_purch
                    records!price = tmpMaterial.getTotalPrice()
                    records.Update
                    records.Close
                    Set records = Nothing
                    Set tmpMaterial = addQuantityMat(tmpMaterial)
                Next i
                CreatePurchaseMaterial = True
            Else
                Debug.Print "Error al asociar compra con materiales: No hay materiales en la compra."
                CreatePurchaseMaterial = False
            End If
            CreatePurchaseMaterial = True
        Else
            Debug.Print "Error al asociar compra con materiales: No se ha creado el proveedor o no hay materiales en la compra."
            CreatePurchaseMaterial = False
        End If
    Else
        Debug.Print "No hay conexion con la base de datos."
        CreatePurchaseMaterial = False
    End If
End Function
