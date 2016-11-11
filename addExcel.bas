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

