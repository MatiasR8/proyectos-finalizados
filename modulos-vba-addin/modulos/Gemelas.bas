Attribute VB_Name = "Gemelas"
Sub BuscarYActualizarTodo()
    ' Macro para buscar valores en la hoja "Gemelas" y actualizar la hoja "LIMS"
    
    ' --- CONFIGURACIÓN INICIAL ---
    ' Optimización: Desactiva actualización de pantalla y cálculos automáticos
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    ' Desprotege la hoja LIMS para permitir modificaciones
    Sheets("LIMS").Unprotect Password:="0000"
    
    ' --- PREPARACIÓN DE VARIABLES ---
    ' Define las hojas de trabajo a utilizar
    Set wsLIMS = ThisWorkbook.Sheets("LIMS")      ' Hoja de destino
    Set wsMuestras = ThisWorkbook.Sheets("Gemelas") ' Hoja de origen
    Set dict = CreateObject("Scripting.Dictionary") ' Diccionario para búsquedas eficientes
    
    ' Limpia la columna R de LIMS donde se escribirán los resultados
    wsLIMS.Columns("R").ClearContents
    
    ' --- CARGA DE DATOS EN DICCIONARIO ---
    ' Recorre las celdas C1:C50 en hoja Gemelas
    For Each celda In wsMuestras.Range("C1:C50")
        ' Si la celda no está vacía, agrega al diccionario:
        ' - Clave: valor de la columna C (celda.Value)
        ' - Valor: valor de la columna B (Offset(0,-1).Value)
        If Not IsEmpty(celda.value) Then
            dict(celda.value) = celda.Offset(0, -1).value
        End If
    Next celda
    
    ' --- BÚSQUEDA Y ACTUALIZACIÓN ---
    ' Define el rango máximo a procesar (fijo en 60000 filas)
    lastRow = 60000
    
    ' Prepara array para almacenar resultados
    ReDim resultados(1 To 60000) ' Array desde posición 1 hasta 60000
    
    ' Recorre cada celda en columna N de LIMS (desde fila 2)
    For i = 2 To lastRow
        valorBuscado = wsLIMS.Cells(i, "N").value
    
        ' Busca el valor en el diccionario
        If dict.Exists(valorBuscado) Then
            ' Si encuentra coincidencia, guarda el valor asociado
            resultados(i - 1) = dict(valorBuscado)
        Else
            ' Si no encuentra coincidencia, guarda cadena vacía
            resultados(i - 1) = ""
        End If
    Next i
    
    ' --- ESCRITURA DE RESULTADOS ---
    ' Transpone el array y escribe los resultados en columna R de LIMS
    ' (Desde R2 hasta R60001)
    wsLIMS.Range("R2").Resize(60000, 1).value = Application.Transpose(resultados)
    
    ' --- FINALIZACIÓN ---
    ' Vuelve a proteger la hoja LIMS
    Sheets("LIMS").Protect Password:="0000"

    ' Restaura configuración inicial de Excel
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
End Sub
