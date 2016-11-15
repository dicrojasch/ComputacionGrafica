Attribute VB_Name = "MyInventor"
'Function IsInventorRunning() As Boolean
'    Dim invApp As Inventor.Application
'    On Error Resume Next
'    Set invApp = GetObject(, "Inventor.Application")
'    IsInventorRunning = (Err.Number = 0)
'    Set invApp = Nothing
'    Err.Clear
'End Function
'
'
''Se debe ingresar como parametro la ruta completa del'
'' archivo de inventor que queremos abrir'
'Public Sub OpenInventorFile(path As String)
'    'Crear un objeto que manejara la aplicacion Inventor'
'    Dim invApp As Inventor.Application
'    Dim inventorRunning As Boolean
'    'Verifica si inventor esta abierto'
'    inventorRunning = IsInventorRunning()
'    'si inventorRunning es true es porque inventor esta abierto'
'    If inventorRunning Then
'        'si inventorRunning es true es porque inventor esta abierto'
'        'Crea un objeto y lo asocia a Inventor y con este objeto se pueden'
'        'hacer varias operaciones sobre la aplicacion'
'        Set invApp = GetObject(, "Inventor.Application")
'    Else
'        'Si inventor no esta abierto, lo abre y despues hace lo mismo que el paso anterior'
'        Set invApp = CreateObject("Inventor.Application")
'        invApp.Visible = True 'Indicamos si que la aplicacion no este oculta'
'
'    End If
'    Do While Not invApp.Ready
'    'Con el while detenemos este script hasta que este inventor este listo'
'    Loop
'    invApp.Documents.CloseAll ' Cerramos todos los documentos que tenga abierto Inventor'
'
'    Dim current As Inventor.Document  'Se crea un documento tipo inventor'
'    Set current = invApp.Documents.Open(path) 'Abre al archivo de inventor'
'    Call ExportTo3D(invApp) 'Exporta un modelo 3D del archivo'
'    current.Close   'Despues de haber exportado el modelo se cierra el documento que abrimos'
'
'    If Not inventorRunning Then
'        invApp.Quit
'        Set invApp = Nothing
'    End If
'End Sub
'
'
'
'Public Sub ExportTo3D(Inv As Inventor.Application)
'    ' Get the 3D PDF Add-In.
'    Dim oPDFAddIn As ApplicationAddIn
'    Dim oAddin As ApplicationAddIn
'    For Each oAddin In Inv.ApplicationAddIns
'        If oAddin.ClassIdString = "{3EE52B28-D6E0-4EA4-8AA6-C2A266DEBB88}" Then
'            Set oPDFAddIn = oAddin
'            Exit For
'        End If
'    Next
'
'    If oPDFAddIn Is Nothing Then
'        MsgBox "Inventor 3D PDF Addin not loaded."
'        Exit Sub
'    End If
'
'    Dim oPDFConvertor3D
'    Set oPDFConvertor3D = oPDFAddIn.Automation
'
'    'Set a reference to the active document (the document to be published).
'    Dim oDocument As Document
'    Set oDocument = Inv.ActiveDocument
'
'    ' Create a NameValueMap object as Options
'    Dim oOptions As NameValueMap
'    Set oOptions = Inv.TransientObjects.CreateNameValueMap
'    DeleteFile (path & modelo3D)
'    ' Options
'    oOptions.value("FileOutputLocation") = path & modelo3D
'    oOptions.value("ExportAnnotations") = 1
'    oOptions.value("ExportWokFeatures") = 1
'    oOptions.value("GenerateAndAttachSTEPFile") = True
'    oOptions.value("VisualizationQuality") = kHigh
'
'    ' Set the properties to export
'    Dim sProps(0) As String
'    sProps(0) = "{F29F85E0-4FF9-1068-AB91-08002B27B3D9}:Title"  ' Title
'
'    oOptions.value("ExportAllProperties") = False
'    oOptions.value("ExportProperties") = sProps
'
'    ' Set the design views to export
'    Dim sDesignViews(1) As String
'    sDesignViews(0) = "Master"
'    sDesignViews(1) = "View1"
'
'    oOptions.value("ExportDesignViewRepresentations") = sDesignViews
'
'    'Publish document.
'    Call oPDFConvertor3D.Publish(oDocument, oOptions)
'End Sub
'
'
'Public Sub ExportTo2D(Inv As Inventor.Application)
'    ' Get the PDF translator Add-In.
'    Dim PDFAddIn As TranslatorAddIn
'    Set PDFAddIn = Inv.ApplicationAddIns.ItemById("{0AC6FD96-2F4D-42CE-8BE0-8AEA580399E4}")
'
'    'Set a reference to the active document (the document to be published).
'    Dim oDocument As Document
'    Set oDocument = Inv.ActiveDocument
'
'    Dim oContext As TranslationContext
'    Set oContext = Inv.TransientObjects.CreateTranslationContext
'    oContext.Type = kFileBrowseIOMechanism
'
'    ' Create a NameValueMap object
'    Dim oOptions As NameValueMap
'    Set oOptions = Inv.TransientObjects.CreateNameValueMap
'
'    ' Create a DataMedium object
'    Dim oDataMedium As DataMedium
'    Set oDataMedium = Inv.TransientObjects.CreateDataMedium
'
'    ' Check whether the translator has 'SaveCopyAs' options
'    If PDFAddIn.HasSaveCopyAsOptions(oDocument, oContext, oOptions) Then
'
'        ' Options for drawings...
'
'        oOptions.value("All_Color_AS_Black") = 0
'
'        'oOptions.Value("Remove_Line_Weights") = 0
'        'oOptions.Value("Vector_Resolution") = 400
'        'oOptions.Value("Sheet_Range") = kPrintAllSheets
'        'oOptions.Value("Custom_Begin_Sheet") = 2
'        'oOptions.Value("Custom_End_Sheet") = 4
'
'    End If
'    DeleteFile (path & modelo2D)
'    'Set the destination file name
'    oDataMedium.FileName = path & modelo2D
'
'    'Publish document.
'    Call PDFAddIn.SaveCopyAs(oDocument, oContext, oOptions, oDataMedium)
'End Sub
'
'
'
'
'
'Public Sub testInventor()
'    Dim test As New CalculateTime
'    test.StartTimer
'    OpenInventorFile (path & pathExample)
'    Debug.Print test.EndTimer
'End Sub
'
'
'
