Attribute VB_Name = "addExcel"

Sub aņadirACelda(ExcelApp As Excel.Application, value As String, position1 As Integer)
    With ExcelApp.ActiveSheet
        If value = "" Or value = "N/A" Then
            .Cells(position1, 5).value = "NO"
        Else
            .Cells(position1, 2).value = value
            .Cells(position1, 5).value = "SI"
        End If
    End With

End Sub

Sub aņadirACelda2(ExcelApp As Excel.Application, value As Boolean, position1 As Integer)
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
    
        Call aņadirACelda(ExcelApp, formaleta.cPlate0, 4)
        Call aņadirACelda(ExcelApp, formaleta.cPlate90, 5)
        Call aņadirACelda(ExcelApp, formaleta.cPlate180, 6)
        Call aņadirACelda(ExcelApp, formaleta.cPlate270, 7)
        Call aņadirACelda(ExcelApp, formaleta.aFPlate0, 8)
        Call aņadirACelda(ExcelApp, formaleta.aFPlate45, 9)
        Call aņadirACelda(ExcelApp, formaleta.aFPlate90, 10)
        Call aņadirACelda(ExcelApp, formaleta.aFPlate135, 11)
        Call aņadirACelda(ExcelApp, formaleta.aFPlate180, 12)
        Call aņadirACelda(ExcelApp, formaleta.aFPlate225, 13)
        Call aņadirACelda(ExcelApp, formaleta.aFPlate270, 14)
        Call aņadirACelda(ExcelApp, formaleta.aFPlate315, 15)

        Call aņadirACelda2(ExcelApp, formaleta.rVar0_90, 16)
        Call aņadirACelda2(ExcelApp, formaleta.rVar90_180, 17)
        Call aņadirACelda2(ExcelApp, formaleta.rVar180_270, 18)
        Call aņadirACelda2(ExcelApp, formaleta.rVar270_0, 19)
        
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

Sub soliMateriales(purchase As Purchases)

    Dim WordApp As Word.Application
    Set WordApp = New Word.Application
    WordApp.Documents.Add (path & "Plantilla Pedir Materiales.dotm")
    
    Call SearchReplace("<<fecha>>", getDate(), WordApp)
    Call stringWord1(WordApp, purchase.toString)

    WordApp.ActiveDocument.SaveAs2 FileName:=path & "Dropbox\Compras\" & "compra" & purchase.id, FileFormat:=wdFormatPDF
    WordApp.ActiveDocument.Saved = True
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
    
    WordApp.ActiveDocument.SaveAs2 FileName:=path & "Dropbox\Compras\" & "cotizacion" & quot.producto.id, FileFormat:=wdFormatPDF
    WordApp.ActiveDocument.Saved = True
    
    WordApp.Quit
    
End Sub
Sub stringWord(WordApp As Word.Application, toString As String)
    
    WordApp.Selection.Move 2, 68
    WordApp.Visible = True
    WordApp.Selection.TypeText (toString)
End Sub
Sub stringWord1(WordApp As Word.Application, toString As String) '
    
    WordApp.Selection.Move 2, 44
    
    WordApp.Visible = True
    
    WordApp.Selection.TypeText (toString)
End Sub
