Attribute VB_Name = "Reportcalibracion"
Sub Calibrarcalibracion()
' Macro para exportar criterios de calibración a archivo PDF
' Maneja diferentes rutas según el método y verifica existencia previa
    
    ' Activa cálculos en hoja Criterios
    Worksheets("Criterios").EnableCalculation = True

    ' --- PREPARACIÓN NOMBRE ARCHIVO ---
    ' Genera nombre PDF con prefijo "Criterios_" + nombre de batch (sin extensión)
    namepdf = "Criterios_" & Split(Sheets("CCD").Range("batch").value, ".")(0)
    
    ' Obtiene ruta base desde hoja Criterios
    ruta = Sheets("Criterios").Range("rutacalibrar").value
    
    ' --- DETERMINACIÓN RUTA FINAL ---
    ' Para métodos especiales usa E13, para otros usa batch normal
    If Sheets("CCD").Range("J12").value <> "CGM/019-pcbbde" And Sheets("CCD").Range("J12").value <> "CGM/031-a-CP" Then
        ' Ruta estándar: reemplaza paréntesis en nombre de batch
        rutacalibrarfinal = ruta & Replace(Replace(Split(Sheets("CCD").Range("batch").value, ".")(0), "(", "-"), ")", "")
    Else
        ' Ruta especial para métodos específicos: usa E13 en lugar de batch
        rutacalibrarfinal = ruta & Replace(Replace(Split(Sheets("CCD").Range("E13").value, ".")(0), "(", "-"), ")", "")
    End If

    ' --- VERIFICACIÓN/CREACIÓN DIRECTORIOS ---
    ' Comprueba existencia de directorio final
    If Dir(rutacalibrarfinal, vbDirectory) = "" Then
        ' Si no existe, verifica primero la ruta base
        If Dir(ruta, vbDirectory) = "" Then
            MsgBox "Hay un error en la ruta de exportación", vbInformation
            GoTo Line1  ' Sale si la ruta base no existe
        End If
        ' Crea directorio específico para este lote
        MkDir rutacalibrarfinal
    End If
    
    ' Marca en hoja Samples que se generó calibración
    Sheets("Samples").Range("AA32").value = "SI"
    
    ' --- VERIFICACIÓN ARCHIVO EXISTENTE ---
    ' Comprueba si ya existe un PDF para este lote
    If Dir(rutacalibrarfinal & "\" & namepdf) <> "" Then
        Dim respuesta As VbMsgBoxResult
        ' Pide confirmación para reemplazar
        respuesta = MsgBox("El archivo '" & namepdf & "' ya existe. ¿Deseas reemplazarlo?", vbYesNo + vbQuestion, "Confirmar Reemplazo")
        If respuesta = vbNo Then
            MsgBox "El archivo no se ha reemplazado.", vbInformation
            GoTo Line1  ' Sale si usuario elige no reemplazar
        End If
    End If
    
    ' --- EXPORTACIÓN A PDF ---
    ' Selecciona hoja Criterios y exporta como PDF
    Sheets("Criterios").Activate
    ChDir rutacalibrarfinal  ' Cambia directorio actual
    ActiveSheet.ExportAsFixedFormat _
        Type:=xlTypePDF, _
        Filename:=rutacalibrarfinal & "\" & namepdf, _
        Quality:=xlQualityStandard, _
        IncludeDocProperties:=False, _
        IgnorePrintAreas:=False, _
        OpenAfterPublish:=False  ' No abrir después de guardar
        
    ' Desactiva cálculos y muestra confirmación
    Worksheets("Criterios").EnableCalculation = False
    MsgBox "La exportación se ha hecho bien", vbInformation
    
Line1:
    ' --- FINALIZACIÓN ---
    ' Vuelve a hoja Samples y desactiva cálculos
    Worksheets("Samples").Activate
    Worksheets("Criterios").EnableCalculation = False
End Sub

