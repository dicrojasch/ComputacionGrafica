Attribute VB_Name = "ExcelFormaleta"

Sub a�adirACelda(excelApp As Excel.Application, value As String, position1 As Integer)
    With excelApp.ActiveSheet
        If value = "" Or value = "N/A" Then
            .Cells(position1, 5).value = "NO"
        Else
            .Cells(position1, 2).value = value
            .Cells(position1, 5).value = "SI"
        End If
    End With

End Sub

Sub a�adirACelda2(excelApp As Excel.Application, value As Boolean, position1 As Integer)
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
    
        Call a�adirACelda(excelApp, formaleta.cPlate0, 4)
        Call a�adirACelda(excelApp, formaleta.cPlate90, 5)
        Call a�adirACelda(excelApp, formaleta.cPlate180, 6)
        Call a�adirACelda(excelApp, formaleta.cPlate270, 7)
        Call a�adirACelda(excelApp, formaleta.aFPlate0, 8)
        Call a�adirACelda(excelApp, formaleta.aFPlate45, 9)
        Call a�adirACelda(excelApp, formaleta.aFPlate90, 10)
        Call a�adirACelda(excelApp, formaleta.aFPlate135, 11)
        Call a�adirACelda(excelApp, formaleta.aFPlate180, 12)
        Call a�adirACelda(excelApp, formaleta.aFPlate225, 13)
        Call a�adirACelda(excelApp, formaleta.aFPlate270, 14)
        Call a�adirACelda(excelApp, formaleta.aFPlate315, 15)

        Call a�adirACelda2(excelApp, formaleta.rVar0_90, 16)
        Call a�adirACelda2(excelApp, formaleta.rVar90_180, 17)
        Call a�adirACelda2(excelApp, formaleta.rVar180_270, 18)
        Call a�adirACelda2(excelApp, formaleta.rVar270_0, 19)
        
    
    excelApp.Visible = True
    excelApp.ActiveWorkbook.Save
    excelApp.ActiveWorkbook.Close
    
End Sub
