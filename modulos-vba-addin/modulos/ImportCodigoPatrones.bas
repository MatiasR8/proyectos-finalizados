Attribute VB_Name = "ImportCodigoPatrones"
Sub ImportTableCodigos()
' Macro para importar c�digos de patrones (QC y CAL) desde un archivo externo
' Selecciona los registros con fechas m�s cercanas a la fecha de referencia

    ' --- CONFIGURACI�N INICIAL ---
    ' Establece hojas de destino
    Set wsDestination = ThisWorkbook.Sheets("Codigos")  ' Hoja principal para datos
    Set wsDestination2 = ThisWorkbook.Sheets("Criterios")  ' Hoja de configuraci�n

    ' Asegura que la hoja Codigos sea visible
    If wsDestination.Visible = xlSheetHidden Then
        wsDestination.Visible = xlSheetVisible
    End If

    ' Desprotege la hoja Codigos para edici�n
    Sheets("Codigos").Unprotect Password:="0000"

    ' --- PREPARACI�N DE ARCHIVO ORIGEN ---
    ' Obtiene ruta del archivo desde celda N5 de Criterios
    sourceWorkbookPath = Sheets("Criterios").Range("N5").value & "Listado Codigos patrones en vigor.xlsx"

    ' Intenta abrir el archivo origen con manejo de errores
    On Error GoTo ErrorHandler
    Set wbSource = Workbooks.Open(sourceWorkbookPath)
    On Error GoTo 0

    ' Obtiene nombre de la hoja origen desde celda O6 de Criterios
    sheetName = wsDestination2.Range("O6").value

    ' Verifica existencia de la hoja origen
    Set wsSource = Nothing
    On Error Resume Next
    Set wsSource = wbSource.Sheets(sheetName)
    On Error GoTo 0

    If wsSource Is Nothing Then
        MsgBox "La hoja '" & sheetName & "' no existe en el libro de origen."
        GoTo Cleanup
    End If

    ' --- PREPARACI�N DE DATOS ---
    ' Obtiene rango completo de datos (tabla contigua desde A1)
    Set tableRange = wsSource.Range("A1").CurrentRegion

    ' Valida fecha de referencia desde celda E8 de Criterios
    If IsDate(wsDestination2.Range("E8").value) Then
        filterDate = wsDestination2.Range("E8").value
    Else
        MsgBox "El nombre del batch no tiene el formato adecuado.", vbCritical, "Error de formato"
        GoTo Cleanup
    End If

    ' Limpia hoja de destino completamente
    wsDestination.Cells.Clear

    ' Copia los encabezados de la tabla origen
    tableRange.Rows(1).Copy Destination:=wsDestination.Rows(1)

    ' --- B�SQUEDA DE FECHAS CERCANAS ---
    ' Inicializa variables para encontrar fechas m�s cercanas
    destinationRow = 2  ' Fila donde empezar a pegar datos
    closestDateQC = filterDate  ' Fecha QC m�s cercana
    closestDateCAL = filterDate ' Fecha CAL m�s cercana
    minDiffQC = 1000   ' Diferencia m�nima para QC (inicializada alta)
    minDiffCAL = 1000  ' Diferencia m�nima para CAL (inicializada alta)
    foundQC = False    ' Bandera si se encontr� QC
    foundCAL = False   ' Bandera si se encontr� CAL

    ' Valores a buscar en columna 6 (QC y CAL)
    filterValues = Array("QC", "CAL")

    ' Busca fechas m�s cercanas para QC y CAL (anteriores a filterDate)
    For Each filterValue In filterValues
        For Each row In tableRange.Rows
            ' Verifica si es fecha v�lida y del tipo buscado (QC/CAL)
            If IsDate(row.Cells(1, 2).value) And row.Cells(1, 6).value = filterValue Then
                diff = filterDate - row.Cells(1, 2).value  ' Diferencia en d�as
                
                ' Actualiza fecha m�s cercana si es anterior y m�s pr�xima
                If diff >= 0 And diff < IIf(filterValue = "QC", minDiffQC, minDiffCAL) Then
                    If filterValue = "QC" Then
                        minDiffQC = diff
                        closestDateQC = row.Cells(1, 2).value
                        foundQC = True
                    ElseIf filterValue = "CAL" Then
                        minDiffCAL = diff
                        closestDateCAL = row.Cells(1, 2).value
                        foundCAL = True
                    End If
                End If
            End If
        Next row
    Next filterValue

    ' --- COPIA DE DATOS SELECCIONADOS ---
    ' Recorre la tabla para copiar registros con fechas encontradas
    For Each row In tableRange.Rows
        If IsDate(row.Cells(1, 2).value) Then
            ' Copia registros QC con fecha exacta encontrada
            If row.Cells(1, 6).value = "QC" And row.Cells(1, 2).value = closestDateQC Then
                row.Copy Destination:=wsDestination.Rows(destinationRow)
                destinationRow = destinationRow + 1
            ' Copia registros CAL con fecha exacta encontrada
            ElseIf row.Cells(1, 6).value = "CAL" And row.Cells(1, 2).value = closestDateCAL Then
                row.Copy Destination:=wsDestination.Rows(destinationRow)
                destinationRow = destinationRow + 1
            End If
        End If
    Next row

    ' --- FINALIZACI�N ---
    ' Muestra resumen de importaci�n con fechas usadas
    MsgBox "Importaci�n completada." & vbNewLine & _
           "Fecha m�s cercana para QC: " & closestDateQC & vbNewLine & _
           "Fecha m�s cercana para CAL: " & closestDateCAL

Cleanup:
    ' Cierra archivo origen sin guardar cambios
    If Not wbSource Is Nothing Then
        wbSource.Close SaveChanges:=False
    End If

    ' Protege hoja Codigos y termina ejecuci�n
    Sheets("Codigos").Protect Password:="0000"
    Exit Sub

ErrorHandler:
    ' Manejo de errores generales
    MsgBox "Ocurri� un error: " & Err.Description
    Resume Cleanup
End Sub
