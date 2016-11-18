Attribute VB_Name = "addExcel"

Sub añadirACelda(excelApp As Excel.Application, value As String, position1 As Integer)
    With excelApp.ActiveSheet
        If value = "" Or value = "N/A" Then
            .Cells(position1, 5).value = "NO"
        Else
            .Cells(position1, 2).value = value
            .Cells(position1, 5).value = "SI"
        End If
    End With

End Sub

Sub añadirACelda2(excelApp As Excel.Application, value As Boolean, position1 As Integer)
    With excelApp.ActiveSheet
         If value Then
            .Cells(position1, 5).value = "SI"
        Else
            .Cells(position1, 5).value = "NO"
        End If
    End With

End Sub

Sub pasarAExcelFormaleta(formaleta As Formaletas, id As Integer)

    Dim excelApp As New Excel.Application
    excelApp.Workbooks.Open (pathFormaletas & "FORMALETAS BASE\DatosEntrada.xlsx") ' archivo excel formaleta

    With excelApp.ActiveSheet
        .Cells(1, 3).value = formaleta.unidades
        .Cells(2, 3).value = formaleta.unidades
        .Cells(3, 3).value = formaleta.unidades
        .Cells(1, 2).value = formaleta.altura
        .Cells(2, 2).value = formaleta.diamInterno
        .Cells(3, 2).value = formaleta.AltRanura
        .Cells(20, 2).value = id
    End With
    
        Call añadirACelda(excelApp, formaleta.cPlate0, 4)
        Call añadirACelda(excelApp, formaleta.cPlate90, 5)
        Call añadirACelda(excelApp, formaleta.cPlate180, 6)
        Call añadirACelda(excelApp, formaleta.cPlate270, 7)
        Call añadirACelda(excelApp, formaleta.aFPlate0, 8)
        Call añadirACelda(excelApp, formaleta.aFPlate45, 9)
        Call añadirACelda(excelApp, formaleta.aFPlate90, 10)
        Call añadirACelda(excelApp, formaleta.aFPlate135, 11)
        Call añadirACelda(excelApp, formaleta.aFPlate180, 12)
        Call añadirACelda(excelApp, formaleta.aFPlate225, 13)
        Call añadirACelda(excelApp, formaleta.aFPlate270, 14)
        Call añadirACelda(excelApp, formaleta.aFPlate315, 15)

        Call añadirACelda2(excelApp, formaleta.rVar0_90, 16)
        Call añadirACelda2(excelApp, formaleta.rVar90_180, 17)
        Call añadirACelda2(excelApp, formaleta.rVar180_270, 18)
        Call añadirACelda2(excelApp, formaleta.rVar270_0, 19)
        
    excelApp.ActiveWorkbook.Save
    excelApp.Quit
    
End Sub
Sub pasarExcelInvernadero(invernadero As Invernaderos)
    Dim excelApp As New Excel.Application
    excelApp.Workbooks.Open (pathInvernaderos & "Parametros_Invernaderos.xlsm") ' archivo excel de invernaderos
    
    With excelApp.ActiveSheet
        .Cells(5, 6).value = invernadero.tipo
        
        .Cells(4, 6).value = invernadero.alto
        
        .Cells(2, 6).value = invernadero.ancho
        
        .Cells(3, 6).value = invernadero.largo
        
    End With
    excelApp.ActiveWorkbook.Save
    
    excelApp.Quit
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
    wordApp.Quit
    
End Sub
Sub wordCotizacion(quot As Quote)

    Dim wordApp As Word.Application
    Set wordApp = New Word.Application
    wordApp.Documents.Add (path & "cotizacion.dotm")
    
    Call SearchReplace("<<date>>", getDate(), wordApp)
    Call SearchReplace("<<clientname>>", quot.cliente.firstName & " " & quot.cliente.lastname, wordApp)
    Call SearchReplace("<<producto>>", quot.producto.getName, wordApp)
    Call stringWord(wordApp, quot.producto.getDescription)
    Call SearchReplace("<<price>>", quot.producto.price, wordApp)
    ' TO DO funcion para calcular product.price
    wordApp.Visible = True
    
    wordApp.ActiveDocument.SaveAs2 FileName:=path & "cotizacion" & quot.producto.id, FileFormat:=wdFormatPDF
    wordApp.ActiveDocument.Saved = True
    
    wordApp.Quit
    
End Sub
Sub stringWord(wordApp As Word.Application, toString As String)
    
    wordApp.Selection.Move 2, 68
    wordApp.Visible = True
    wordApp.Selection.TypeText (toString)
End Sub
