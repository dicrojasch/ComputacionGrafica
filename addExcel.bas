Attribute VB_Name = "addExcel"

Sub aņadirACelda(excelApp As Excel.Application, value As String, position1 As Integer)
    With excelApp.ActiveSheet
        If value = "" Or value = "N/A" Then
            .Cells(position1, 5).value = "NO"
        Else
            .Cells(position1, 2).value = value
            .Cells(position1, 5).value = "SI"
        End If
    End With

End Sub

Sub aņadirACelda2(excelApp As Excel.Application, value As Boolean, position1 As Integer)
    With excelApp.ActiveSheet
         If value Then
            .Cells(position1, 5).value = "SI"
        Else
            .Cells(position1, 5).value = "NO"
        End If
    End With

End Sub

Sub pasarAExcelFormaleta(formaleta As Formaletas)

    Dim excelApp As New Excel.Application
    excelApp.Workbooks.Open (path & "Plantilla de datos.xlsx") ' archivo excel formalera

    With excelApp.ActiveSheet
        .Cells(1, 3).value = formaleta.unidades
        .Cells(2, 3).value = formaleta.unidades
        .Cells(3, 3).value = formaleta.unidades
        .Cells(1, 2).value = formaleta.altura
        .Cells(2, 2).value = formaleta.diamInterno
        .Cells(3, 2).value = formaleta.AltRanura
        
    End With
    
        Call aņadirACelda(excelApp, formaleta.cPlate0, 4)
        Call aņadirACelda(excelApp, formaleta.cPlate90, 5)
        Call aņadirACelda(excelApp, formaleta.cPlate180, 6)
        Call aņadirACelda(excelApp, formaleta.cPlate270, 7)
        Call aņadirACelda(excelApp, formaleta.aFPlate0, 8)
        Call aņadirACelda(excelApp, formaleta.aFPlate45, 9)
        Call aņadirACelda(excelApp, formaleta.aFPlate90, 10)
        Call aņadirACelda(excelApp, formaleta.aFPlate135, 11)
        Call aņadirACelda(excelApp, formaleta.aFPlate180, 12)
        Call aņadirACelda(excelApp, formaleta.aFPlate225, 13)
        Call aņadirACelda(excelApp, formaleta.aFPlate270, 14)
        Call aņadirACelda(excelApp, formaleta.aFPlate315, 15)

        Call aņadirACelda2(excelApp, formaleta.rVar0_90, 16)
        Call aņadirACelda2(excelApp, formaleta.rVar90_180, 17)
        Call aņadirACelda2(excelApp, formaleta.rVar180_270, 18)
        Call aņadirACelda2(excelApp, formaleta.rVar270_0, 19)
        
    
    excelApp.Visible = True
    excelApp.ActiveWorkbook.Save
    excelApp.ActiveWorkbook.Close
    
End Sub
Sub pasarExcelInvernadero(invernadero As Invernaderos)
    Dim excelApp As New Excel.Application
    excelApp.Workbooks.Open (path & "Parametros_Invernaderos.xlsm") ' archivo excel de invernaderos
    
    With excelApp.ActiveSheet
        .Cells(5, 6).value = invernadero.tipo
        
        .Cells(4, 6).value = invernadero.alto
        
        .Cells(2, 6).value = invernadero.ancho
        
        .Cells(3, 6).value = invernadero.largo
        
    End With
    
    excelApp.Visible = True
    excelApp.ActiveWorkbook.Save
    excelApp.ActiveWorkbook.Close
End Sub

Private Sub SearchReplace(search As String, replace As String, wordApp As Word.Application)
    Dim FindObject As Word.Find
    Set FindObject = wordApp.Selection.Find
    With FindObject
        .ClearFormatting
        .Text = search
        .Replacement.ClearFormatting
        .Replacement.Text = replace
    End With
    Call FindObject.Execute(replace:=Word.WdReplace.wdReplaceAll)
End Sub

Sub ShowSelection(purchase As Purchases)

    Dim wordApp As Word.Application
    Set wordApp = New Word.Application
    wordApp.Documents.Open (path & "Plantilla Pedir Materiales.dotm")
    
    Call SearchReplace("<<fecha>>", getDate(), wordApp)
    Call SearchReplace("<<proveedor>>", purchase.provider_name, wordApp)
    Call SearchReplace("<<materiales>>", purchase.toString, wordApp)
    
    wordApp.ActiveDocument.SaveAs2 (path & "compra" & purchase.id & ".docx")
    wordApp.ActiveDocument.Close
    
End Sub
Sub wordCotizacion(quot As Quote)

    Dim wordApp As Word.Application
    Set wordApp = New Word.Application
    wordApp.Documents.Open (path & "cotizacion.dotm")
    
    Call SearchReplace("<<date>>", getDate(), wordApp)
    Call SearchReplace("<<clientname>>", quot.cliente.firstName & " " & quot.cliente.lastname, wordApp)
    Call SearchReplace("<<producto>>", quot.producto.getName, wordApp)
    Call SearchReplace("<<parameters>>", quot.producto.toString, wordApp)
    Call SearchReplace("<<price>>", quot.producto.price, wordApp)
    ' TO DO funcion para calcular product.price
    wordApp.ActiveDocument.SaveAs2 (path & "cotizacion" & quot.producto.id & ".docx")
    wordApp.ActiveDocument.Close
    
End Sub
