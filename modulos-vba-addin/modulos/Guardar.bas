Attribute VB_Name = "Guardar"
Sub Guardar_datos()
' Macro para guardar muestras desde la hoja "Sample" con validaciones y exportaciones

    ' --- VALIDACI�N DE BLANCOS ---
    ' Verifica si hay que comprobar blancos (AF32 = "SI")
    If Sheets("Samples").Range("AF32").value = "SI" Then
        Dim cell As Range
        Dim continuar As Boolean
        continuar = True
        
        ' Revisa el rango AB40:AB190 para valores problem�ticos en blancos
        For Each cell In Sheets("Samples").Range("AB40:AB190")
            ' Detecta valores "NOT OK", "FR" o "POS" en muestras marcadas como blancos (col M = TRUE)
            If (cell.value = "NOT OK" Or cell.value = "FR" Or cell.value = "POS") And Sheets("Samples").Range("M" & cell.row).value = True Then
                ' Pide confirmaci�n al usuario
                respuesta = MsgBox("Est�s intentando subir un blanco con valores que no son negativos, �Deseas continuar?", vbYesNo + vbExclamation, "Confirmar acci�n")
                
                If respuesta = vbNo Then
                    continuar = False
                    Exit For
                End If
                Exit For
            End If
        Next cell
        
        ' Si el usuario cancela, termina la ejecuci�n
        If Not continuar Then Exit Sub
    End If

    ' --- EXPORTACI�N DE REPORTE ---
    ' Paso 1: Exporta el reporte antes de cualquier modificaci�n
    Call ExportReport
    
    ' --- EXPORTACI�N DE DATOS PARA LIMS ---
    ' Paso 2: Exporta datos seg�n el m�todo elegido por el usuario
    If Sheets("Samples").Range("U32").value = "No" Then
        Call normalguardar  ' Exportaci�n est�ndar
    ElseIf Sheets("Samples").Range("U32").value = "Yes" Then
        Call selectionguardar  ' Exportaci�n selectiva
    End If
    
    ' --- VALIDACIONES FINALES ---
    ' Asegura que ISTD y QC est�n marcados como "Yes"
    If Sheets("Samples").Range("N32").value <> "Yes" Then
        Sheets("Samples").Range("N32").value = "Yes"
    End If
    If Sheets("Samples").Range("Q32").value <> "Yes" Then
         Sheets("Samples").Range("Q32").value = "Yes"
    End If
    
    ' Mensajes recordatorios para el usuario
    If Sheets("Samples").Range("M341").value = True Then
        MsgBox "Recuerda a�adir los par�metros en LIMS"
    End If
    
    If Sheets("Gemelas").Range("X9").value = False Then
        MsgBox "Esta muestra tiene el par�metro " & Sheets("Gemelas").Range("X10").value
    End If
    
    If Sheets("Gemelas").Range("X16").value = "SI" Then
        MsgBox "Esta muestra tiene el par�metro 3659/3660"
    End If
    
End Sub

Sub normalguardar()
' Exportaci�n est�ndar cuando la muestra aparece una vez en el batch

    ' Obtiene c�digo de muestra desde F26 (considerando dos formatos posibles)
    On Error Resume Next
    muestra = Left(Sheets("Samples").Range("F26").value, InStr(Sheets("Samples").Range("F26").value, " ") - 1)
    If muestra = "" Then muestra = Left(Sheets("Samples").Range("F26"), 6)
    On Error GoTo 0
    
    ' Prepara hojas de trabajo
    Set wsOrigen = Sheets("Informefinal")
    wsOrigen.Unprotect Password:="0000"
    Set wsDestino = Worksheets("Exportacion")
    Set wsDestino2 = Worksheets("Par�metros_Barrido")
    wsDestino.Unprotect Password:="0000"
    wsDestino2.Unprotect Password:="0000"
    Sheets("Samples").Select
    
    ' Verifica si hay datos para exportar
    If Sheets("InformeFinal").Range("K1").value = 0 Then
        MsgBox "No hay compuestos pendientes en el lims a guardar", vbInformation
        GoTo Line1
    End If
    
    ' --- MANEJO DE DATOS PREVIOS ---
    ' Si la muestra ya fue exportada, pregunta si reemplazar
    If ActiveSheet.Range("I32").value = "Exportado" Then
        result = MsgBox("�Quieres reemplazar los datos a exportar?", vbOKCancel)
        If result = vbCancel Then
                GoTo Line1
        ElseIf result = vbOK Then
                ' Elimina datos previos de esta muestra
                wsDestino.Select
                ActiveSheet.Unprotect
                Rows("1:1").Insert Shift:=xlDown
                Set miRango = wsDestino.Range("A1:J20000")
                miRango.AutoFilter Field:=2, Criteria1:="*" & muestra & "*"
                miRango.EntireRow.Delete
        End If
    End If
    
    ' --- EXPORTACI�N PRINCIPAL ---
    Application.ScreenUpdating = False
    wsOrigen.Select
    
    ' Selecciona rango de datos desde A3:N3 hasta final
    Range("A3:N3").Select
    Range(Selection, Selection.End(xlDown)).Select
    Set finalrange = Selection
    
    ' Filtra solo datos marcados para exportar (columna K = 1)
    finalrange.AutoFilter Field:=11, Criteria1:="1"

    ' Encuentra �ltima fila vac�a en hoja de exportaci�n
    ultimaFila = wsDestino.Cells(Rows.Count, 1).End(xlUp).row
    If ultimaFila = 1 And IsEmpty(wsDestino.Range("A1").value) Then ultimaFila = 0
    
    ' Copia datos visibles (filtrados) a hoja de exportaci�n
    Set exportrango = Worksheets("InformeFinal").Range("A4:J149").SpecialCells(xlCellTypeVisible)
    exportrango.Copy
    wsDestino.Cells(ultimaFila + 1, 1).PasteSpecial Paste:=xlPasteValues
    finalrange.AutoFilter
    
    ' --- MANEJO DE PAR�METROS DE BARRIDO ---
    ' Si hay par�metros de barrido (M341 = TRUE)
    If Sheets("Samples").Range("M341").value = True Then
        wsOrigen.Select
        ' Procesa rango de par�metros de barrido (AR1:BB1 hasta final)
        Range("AR1:BB1").Select
        Range(Selection, Selection.End(xlDown)).Select
        Set finalrange2 = Selection
        finalrange2.AutoFilter Field:=11, Criteria1:="1"
        
        ' Exporta a hoja principal
        ultimaFila = wsDestino.Cells(Rows.Count, 1).End(xlUp).row
        If ultimaFila = 1 And IsEmpty(wsDestino.Range("A1").value) Then ultimaFila = 0
        
        ' Exporta a hoja de par�metros
        ultimaFila2 = wsDestino2.Cells(Rows.Count, 1).End(xlUp).row
        If ultimaFila2 = 1 And IsEmpty(wsDestino2.Range("A1").value) Then ultimaFila2 = 0
        
        ' Copia datos con manejo de errores
        On Error Resume Next
        Set exportrango2 = Worksheets("InformeFinal").Range("AR2:BA150").SpecialCells(xlCellTypeVisible)
        exportrango2.Copy
        wsDestino.Cells(ultimaFila + 1, 1).PasteSpecial Paste:=xlPasteValues
        
        ' Copia completa a hoja de par�metros
        Set exportrango2 = Worksheets("InformeFinal").Range("AR2:BA150")
        exportrango2.Copy
        wsDestino2.Cells(ultimaFila2 + 1, 1).PasteSpecial Paste:=xlPasteValues
        
        ' Limpia filtros
        finalrange2.AutoFilter
        finalrange2.Parent.AutoFilterMode = False
        
        ' Elimina filas vac�as en hoja de par�metros
        wsDestino2.Range("A1:A" & ultimaFila2).AutoFilter Field:=1, Criteria1:="="
        On Error Resume Next
        Set rango = wsDestino2.Range("A2:A" & ultimaFila2).SpecialCells(xlCellTypeVisible)
        On Error GoTo 0
        If Not rango Is Nothing Then rango.EntireRow.Delete
        wsDestino2.AutoFilterMode = False
    End If
            
Line1:
    ' --- FINALIZACI�N ---
    ' Vuelve a proteger hojas y restaura configuraci�n
    wsOrigen.Protect Password:="0000"
    wsDestino.Protect Password:="0000"
    wsDestino2.Protect Password:="0000"
    Sheets("Samples").Select
    Application.ScreenUpdating = False
End Sub

Sub selectionguardar()
' Exportaci�n selectiva cuando la muestra aparece dos veces en el batch

    ' Obtiene c�digo de muestra desde F26
    On Error Resume Next
    muestra = Left(Sheets("Samples").Range("F26").value, InStr(Sheets("Samples").Range("F26").value, " ") - 1)
    If muestra = "" Then muestra = Left(Sheets("Samples").Range("F26"), 6)
    On Error GoTo 0
    
    ' Prepara hojas de trabajo
    Set wsOrigen = Sheets("Informefinal")
    wsOrigen.Unprotect Password:="0000"
    Set wsDestino = Worksheets("Exportacion")
    wsDestino.Unprotect Password:="0000"
    Sheets("Samples").Select
    
    ' Verifica si hay datos para exportar
    If Sheets("InformeFinal").Range("L1").value = 0 Then
        MsgBox "No hay compuestos pendientes en el lims a guardar", vbInformation
        GoTo Line1
    End If
    
    ' --- EXPORTACI�N PRINCIPAL ---
    Application.ScreenUpdating = False
    wsOrigen.Select
    
    ' Selecciona rango de datos desde A3:N3 hasta final
    Range("A3:N3").Select
    Range(Selection, Selection.End(xlDown)).Select
    Set finalrange = Selection
    
    ' Filtra datos marcados para exportar (col K=1) y seleccionados (col N=1)
    With finalrange
        .AutoFilter Field:=11, Criteria1:="1"  ' Datos para importar
        .AutoFilter Field:=14, Criteria1:="1"  ' Datos seleccionados
    End With

    ' Encuentra �ltima fila vac�a y copia datos
    ultimaFila = wsDestino.Cells(Rows.Count, 1).End(xlUp).row
    Set exportrango = Worksheets("InformeFinal").Range("A4:J150").SpecialCells(xlCellTypeVisible)
    exportrango.Copy
    If ultimaFila = 1 And IsEmpty(wsDestino.Range("A1").value) Then ultimaFila = 0
    wsDestino.Cells(ultimaFila + 1, 1).PasteSpecial Paste:=xlPasteValues
    finalrange.AutoFilter
    
    ' --- RESET DE SELECCIONES ---
    Sheets("Samples").Activate
    Range("U32").value = "No"  ' Vuelve a modo normal
    Range("AF40:AF341").ClearContents  ' Limpia selecciones
    Range("I26").Select
    
Line1:
    ' --- FINALIZACI�N ---
    Range("U32").value = "No"  ' Asegura reset
    wsOrigen.Protect Password:="0000"
    wsDestino.Protect Password:="0000"
    Sheets("Samples").Select
    Application.ScreenUpdating = False
End Sub
