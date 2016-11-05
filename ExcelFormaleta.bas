Attribute VB_Name = "ExcelFormaleta"

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

Sub pasarAExcelFormaleta(formaleta As Formaletas)

    Dim excelApp As New Excel.Application
    excelApp.Workbooks.Open (path & "Plantilla de datos.xlsx")

    With excelApp.ActiveSheet
        .Cells(1, 3).value = formaleta.unidades
        .Cells(2, 3).value = formaleta.unidades
        .Cells(3, 3).value = formaleta.unidades
        .Cells(1, 2).value = formaleta.altura
        .Cells(2, 2).value = formaleta.diamInterno
        .Cells(3, 2).value = formaleta.AltRanura
        
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
        
    
    excelApp.Visible = True
    excelApp.ActiveWorkbook.Save
    excelApp.ActiveWorkbook.Close
    
End Sub
