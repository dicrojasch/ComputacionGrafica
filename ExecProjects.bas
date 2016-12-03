Attribute VB_Name = "ExecProjects"
'Execute projects
Sub ExecInvernaderos()
    Dim ExcelApp As Excel.Application
    Set ExcelApp = New Excel.Application
    ExcelApp.Workbooks.Open (pathInvernaderos & "Parametros_Invernaderos.xlsm") ' archivo excel de invernaderos
    
    ExcelApp.Run "SEnsamble_Piramide"                                           'ejecuta la macro de invernaderos
    Do While Dir(pathInvernaderos & "Cotizacion_BT4.pdf") = "" _
                And Dir(pathInvernaderos & "Cotizacion_BT7.pdf") = "" _
                And Dir(pathInvernaderos & "Cotizacion_RC.pdf") = "" _
                And Dir(pathInvernaderos & "Cotizacion_RT.pdf") = ""
    Loop
    ExcelApp.ActiveWorkbook.Save
    ExcelApp.Quit
    
    Call closeExcel
    Call closeWord
    
End Sub
Sub ExecFormaletas()

    Call closeExcel
    
    Dim ExcelApp As Excel.Application
    Set ExcelApp = New Excel.Application
    ExcelApp.Workbooks.Open (pathFormaletas & "MacroBonita.xlsm") 'path macro bonita
    Do While Dir(pathInvernaderos & "Plano 2.pdf") = ""
    Loop
    PauseTime = 20    ' Set duration.
    Start = Timer    ' Set start time.
    Do While Timer < Start + PauseTime
    Loop
    
    Call closeInventor
    
    Call closeExcel
    
    PauseTime = 5   ' Set duration.
    Start = Timer    ' Set start time.
    Do While Timer < Start + PauseTime
    Loop
    Call ExecCoti
End Sub
Sub ExecCoti()
    Dim ExcelApp As Excel.Application
    Set ExcelApp = New Excel.Application
    ExcelApp.Workbooks.Open (pathFormaletas & "inventario.xlsm")
    Do While Dir(pathInvernaderos & "Cotizacion_Formaleta.pdf") = ""
    Loop
    ExcelApp.ActiveWorkbook.Close (0)
    ExcelApp.Quit
End Sub


