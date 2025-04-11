Attribute VB_Name = "ImportBlancos_C40_Gemela"
Sub Blancos()
    ' Macro para importar datos de blancos y gemelas desde archivos externos
    ' según el método de análisis utilizado

    ' --- SELECCIÓN DE ARCHIVO SEGÚN MÉTODO ---
    ' Verifica si el método NO es CGM/031 ni CGM/019
    If Left(Sheets("CCD").Range("J12").value, 7) <> "CGM/031" And Left(Sheets("CCD").Range("J12").value, 7) <> "CGM/019" Then
        
        ' --- PRIMER ESCENARIO (métodos normales) ---
        ' Abre archivo de gestión de blancos y gemelas
        Set wbDestino = ThisWorkbook
        Set wbOrigen = Workbooks.Open("\\ruta-censurada\8. LABORATORIO CROMATOGRAFIA CS\23.GESTIÓN Y CONTROL\23.1.GESTIÓN Y CONTROL-SOIL-LAB\23.1.4.CONTROL SOIL-LAB\URGENCIAS-BLANCOS-REVISAR-GEMELAS.xlsx")
        
        ' Configura hojas de trabajo
        Set wsBlancos = wbOrigen.Sheets("Blancos")
        Set wsGemelas = wbDestino.Sheets("Gemelas")
        Set wsGemelaOrigen = wbOrigen.Sheets("Gemelas")
        
        ' Prepara hoja Gemelas para edición
        wsGemelas.Unprotect Password:="0000"
        
        ' Limpia datos existentes
        lastRow = wsGemelas.Cells(wsGemelas.Rows.Count, "B").End(xlUp).row
        wsGemelas.Range("B2:C" & lastRow).ClearContents
        lastRow = wsGemelas.Cells(wsGemelas.Rows.Count, "T").End(xlUp).row
        wsGemelas.Range("T2:T" & lastRow).ClearContents
        
        ' Obtiene rangos de datos origen
        ultimaFila = wsBlancos.Cells(wsBlancos.Rows.Count, "B").End(xlUp).row
        ultimaFila2 = wsGemelaOrigen.Cells(wsGemelaOrigen.Rows.Count, "B").End(xlUp).row
        
        ' Define rangos a copiar
        Set rangoOrigen = wsBlancos.Range("B4:B" & ultimaFila)         ' Blancos
        Set rangoGemela = wsGemelaOrigen.Range("B4:B" & ultimaFila2)   ' Gemelas (col B)
        Set rangoGemela2 = wsGemelaOrigen.Range("C4:C" & ultimaFila2)  ' Gemelas (col C)
        
        ' Copia datos a hoja Gemelas del libro actual
        rangoOrigen.Copy
        wsGemelas.Range("T2").PasteSpecial Paste:=xlPasteValues
        rangoGemela.Copy
        wsGemelas.Range("B2").PasteSpecial Paste:=xlPasteValues
        rangoGemela2.Copy
        wsGemelas.Range("C2").PasteSpecial Paste:=xlPasteValues
        
        ' Cierra archivo origen sin guardar
        wbOrigen.Close SaveChanges:=False
        wsGemelas.Protect Password:="0000"
    Else
        ' --- SEGUNDO ESCENARIO (métodos CGM/031 o CGM/019) ---
        ' Abre archivo específico para semivolátiles
        Set wbDestino = ThisWorkbook
        Set wbOrigen = Workbooks.Open("\\ruta-censurada\8. LABORATORIO CROMATOGRAFIA CS\6. LIBRETAS Y PLANTILLAS\4.Libretas digitales\Semivolatiles\Gemelas.xlsx")
        
        ' Configura hojas de trabajo (similar al primer escenario)
        Set wsBlancos = wbOrigen.Sheets("BLANCOS")
        Set wsGemelas = wbDestino.Sheets("Gemelas")
        Set wsGemelaOrigen = wbOrigen.Sheets("Gemelas")
        
        ' Resto del proceso es idéntico al primer escenario
        wsGemelas.Unprotect Password:="0000"
        
        lastRow = wsGemelas.Cells(wsGemelas.Rows.Count, "B").End(xlUp).row
        wsGemelas.Range("B2:C" & lastRow).ClearContents
        lastRow = wsGemelas.Cells(wsGemelas.Rows.Count, "T").End(xlUp).row
        wsGemelas.Range("T2:T" & lastRow).ClearContents
        
        ultimaFila = wsBlancos.Cells(wsBlancos.Rows.Count, "B").End(xlUp).row
        ultimaFila2 = wsGemelaOrigen.Cells(wsGemelaOrigen.Rows.Count, "B").End(xlUp).row

        Set rangoOrigen = wsBlancos.Range("B4:B" & ultimaFila)
        Set rangoGemela = wsGemelaOrigen.Range("B4:B" & ultimaFila2)
        Set rangoGemela2 = wsGemelaOrigen.Range("C4:C" & ultimaFila2)
        
        rangoOrigen.Copy
        wsGemelas.Range("T2").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, Skipblanks _
        :=False, Transpose:=False
        rangoGemela.Copy
        wsGemelas.Range("B2").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, Skipblanks _
        :=False, Transpose:=False
        rangoGemela2.Copy
        wsGemelas.Range("C2").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, Skipblanks _
        :=False, Transpose:=False
        
        wbOrigen.Close SaveChanges:=False
        
        wsGemelas.Protect Password:="0000"
    End If
    
    Application.CutCopyMode = False
End Sub

Sub c5c40()
    ' Macro para importar datos de control de muestras C5-C40 a la hoja Gemelas

    ' Abre archivo de control de muestras (sin actualizar vínculos)
    Workbooks.Open Filename:="\\ruta-censurada\8. LABORATORIO CROMATOGRAFIA CS\23.GESTIÓN Y CONTROL\23.1.GESTIÓN Y CONTROL-SOIL-LAB\23.1.1.LIBRETAS CONTROL DE MUESTRAS\C5-C40_Control muestras.xlsm", UpdateLinks:=0
    
    Set WB = Workbooks.Open("\\ruta-censurada\8. LABORATORIO CROMATOGRAFIA CS\23.GESTIÓN Y CONTROL\23.1.GESTIÓN Y CONTROL-SOIL-LAB\23.1.1.LIBRETAS CONTROL DE MUESTRAS\C5-C40_Control muestras.xlsm")
    Set ws = WB.Sheets(3)  ' Usa la tercera hoja
    
    ' Prepara hoja origen
    ws.Unprotect
    On Error Resume Next
    ws.ShowAllData  ' Elimina filtros existentes
    On Error GoTo 0
    
    ' Configura filtro para valores vacíos en columna O (15ª columna del rango B:P)
    Set filterRange = ws.Range("B4:P4")
    filterRange.AutoFilter Field:=15, Criteria1:="="
    
    ' Encuentra última fila con datos en columnas B y F
    lastRowB = ws.Cells(ws.Rows.Count, "B").End(xlUp).row
    lastRowP = ws.Cells(ws.Rows.Count, "F").End(xlUp).row
    lastRow = Application.WorksheetFunction.Max(lastRowB, lastRowP)
    
    ' Obtiene rangos visibles (filtrados)
    Set rngB = ws.Range("B5:B" & lastRow).SpecialCells(xlCellTypeVisible)  ' Códigos
    Set rngP = ws.Range("F5:F" & lastRow).SpecialCells(xlCellTypeVisible)  ' Parámetros
    
    ' Prepara hoja Gemelas en libro actual
    Set gemelasWs = ThisWorkbook.Sheets("Gemelas")
    gemelasWs.Unprotect Password:="0000"
    
    ' Limpia datos existentes
    gemelasWs.Range("V2:V" & gemelasWs.Cells(gemelasWs.Rows.Count, "V").End(xlUp).row).ClearContents
    gemelasWs.Range("W2:W" & gemelasWs.Cells(gemelasWs.Rows.Count, "W").End(xlUp).row).ClearContents
    
    ' Copia datos filtrados
    rngB.Copy
    gemelasWs.Range("V2").PasteSpecial Paste:=xlPasteValues  ' Códigos en col V
    rngP.Copy
    gemelasWs.Range("W2").PasteSpecial Paste:=xlPasteValues  ' Parámetros en col W

    ' Cierra archivo origen sin guardar
    WB.Close SaveChanges:=False
    
    ' Protege hoja Gemelas
    gemelasWs.Protect Password:="0000"
End Sub

