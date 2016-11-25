Attribute VB_Name = "addExcel"

Sub añadirACelda(ExcelApp As Excel.Application, value As String, position1 As Integer)
    With ExcelApp.ActiveSheet
        If value = "" Or value = "N/A" Then
            .Cells(position1, 5).value = "NO"
        Else
            .Cells(position1, 2).value = value
            .Cells(position1, 5).value = "SI"
        End If
    End With

End Sub

Sub añadirACelda2(ExcelApp As Excel.Application, value As Boolean, position1 As Integer)
    With ExcelApp.ActiveSheet
         If value Then
            .Cells(position1, 5).value = "SI"
        Else
            .Cells(position1, 5).value = "NO"
        End If
    End With

End Sub

Sub pasarAExcelFormaleta(formaleta As Formaletas, id As Integer)

    Dim ExcelApp As New Excel.Application
    ExcelApp.Workbooks.Open (pathFormaletas & "DatosEntrada.xlsx") ' archivo excel formaleta

    With ExcelApp.ActiveSheet
        .Cells(1, 3).value = formaleta.unidades
        .Cells(2, 3).value = formaleta.unidades
        .Cells(3, 3).value = formaleta.unidades
        .Cells(1, 2).value = formaleta.altura
        .Cells(2, 2).value = formaleta.diamInterno
        .Cells(3, 2).value = formaleta.AltRanura
        .Cells(20, 2).value = id
    End With
    
        Call añadirACelda(ExcelApp, formaleta.cPlate0, 4)
        Call añadirACelda(ExcelApp, formaleta.cPlate90, 5)
        Call añadirACelda(ExcelApp, formaleta.cPlate180, 6)
        Call añadirACelda(ExcelApp, formaleta.cPlate270, 7)
        Call añadirACelda(ExcelApp, formaleta.aFPlate0, 8)
        Call añadirACelda(ExcelApp, formaleta.aFPlate45, 9)
        Call añadirACelda(ExcelApp, formaleta.aFPlate90, 10)
        Call añadirACelda(ExcelApp, formaleta.aFPlate135, 11)
        Call añadirACelda(ExcelApp, formaleta.aFPlate180, 12)
        Call añadirACelda(ExcelApp, formaleta.aFPlate225, 13)
        Call añadirACelda(ExcelApp, formaleta.aFPlate270, 14)
        Call añadirACelda(ExcelApp, formaleta.aFPlate315, 15)

        Call añadirACelda2(ExcelApp, formaleta.rVar0_90, 16)
        Call añadirACelda2(ExcelApp, formaleta.rVar90_180, 17)
        Call añadirACelda2(ExcelApp, formaleta.rVar180_270, 18)
        Call añadirACelda2(ExcelApp, formaleta.rVar270_0, 19)
        
    ExcelApp.ActiveWorkbook.Save
    ExcelApp.Quit
    
End Sub
Sub pasarExcelInvernadero(invernadero As Invernaderos)
    Dim ExcelApp As Excel.Application
    Set ExcelApp = New Excel.Application
    
    ExcelApp.Workbooks.Open (pathInvernaderos & "Parametros_Invernaderos.xlsm") ' archivo excel de invernaderos
    
    With ExcelApp.ActiveSheet
        
        .Cells(5, 6).value = invernadero.tipo
        
        .Cells(4, 6).value = invernadero.alto
        
        .Cells(2, 6).value = invernadero.ancho
        
        .Cells(3, 6).value = invernadero.largo
        
    End With
    ExcelApp.ActiveWorkbook.Save
    
    ExcelApp.Quit
End Sub

Private Sub SearchReplace(search As String, replace As String, WordApp As Word.Application)
    Dim FindObject As Word.Find
    Set FindObject = WordApp.Selection.Find
    With FindObject
        .ClearFormatting
        .Text = search
        .Replacement.ClearFormatting
        .Replacement.Text = replace
        
    End With
    Call FindObject.Execute(replace:=Word.WdReplace.wdReplaceAll)
End Sub

Sub ShowSelection(purchase As Purchases)

    Dim WordApp As Word.Application
    Set WordApp = New Word.Application
    WordApp.Documents.Open (path & "Plantilla Pedir Materiales.dotm")
    
    Call SearchReplace("<<fecha>>", getDate(), WordApp)
    Call SearchReplace("<<proveedor>>", purchase.provider_name, WordApp)
    Call SearchReplace("<<materiales>>", purchase.toString, WordApp)
    
    WordApp.ActiveDocument.SaveAs2 (path & "compra" & purchase.id & ".docx")
    WordApp.Quit
    
End Sub
Sub wordCotizacion(quot As Quote)

    Dim WordApp As Word.Application
    Set WordApp = New Word.Application
    WordApp.Documents.Add (path & "cotizacion.dotm")
    
    Call SearchReplace("<<date>>", getDate(), WordApp)
    Call SearchReplace("<<clientname>>", quot.cliente.firstName & " " & quot.cliente.lastname, WordApp)
    Call SearchReplace("<<producto>>", quot.producto.getName, WordApp)
    Call stringWord(WordApp, quot.producto.getDescription)
    Call SearchReplace("<<price>>", quot.producto.price, WordApp)
    ' TO DO funcion para calcular product.price
    WordApp.Visible = True
    
    WordApp.ActiveDocument.SaveAs2 FileName:=path & "cotizacion" & quot.producto.id, FileFormat:=wdFormatPDF
    WordApp.ActiveDocument.Saved = True
    
    WordApp.Quit
    
End Sub
Sub stringWord(WordApp As Word.Application, toString As String)
    
    WordApp.Selection.Move 2, 68
    WordApp.Visible = True
    WordApp.Selection.TypeText (toString)
End Sub
