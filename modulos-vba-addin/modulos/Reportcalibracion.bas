Attribute VB_Name = "Reportcalibracion"
Sub Calibrarcalibracion()
' Macro para exportar criterios de calibraci�n a archivo PDF
' Maneja diferentes rutas seg�n el m�todo y verifica existencia previa
    
    ' Activa c�lculos en hoja Criterios
    Worksheets("Criterios").EnableCalculation = True

    ' --- PREPARACI�N NOMBRE ARCHIVO ---
    ' Genera nombre PDF con prefijo "Criterios_" + nombre de batch (sin extensi�n)
    namepdf = "Criterios_" & Split(Sheets("CCD").Range("batch").value, ".")(0)
    
    ' Obtiene ruta base desde hoja Criterios
    ruta = Sheets("Criterios").Range("rutacalibrar").value
    
    ' --- DETERMINACI�N RUTA FINAL ---
    ' Para m�todos especiales usa E13, para otros usa batch normal
    If Sheets("CCD").Range("J12").value <> "CGM/019-pcbbde" And Sheets("CCD").Range("J12").value <> "CGM/031-a-CP" Then
        ' Ruta est�ndar: reemplaza par�ntesis en nombre de batch
        rutacalibrarfinal = ruta & Replace(Replace(Split(Sheets("CCD").Range("batch").value, ".")(0), "(", "-"), ")", "")
    Else
        ' Ruta especial para m�todos espec�ficos: usa E13 en lugar de batch
        rutacalibrarfinal = ruta & Replace(Replace(Split(Sheets("CCD").Range("E13").value, ".")(0), "(", "-"), ")", "")
    End If

    ' --- VERIFICACI�N/CREACI�N DIRECTORIOS ---
    ' Comprueba existencia de directorio final
    If Dir(rutacalibrarfinal, vbDirectory) = "" Then
        ' Si no existe, verifica primero la ruta base
        If Dir(ruta, vbDirectory) = "" Then
            MsgBox "Hay un error en la ruta de exportaci�n", vbInformation
            GoTo Line1  ' Sale si la ruta base no existe
        End If
        ' Crea directorio espec�fico para este lote
        MkDir rutacalibrarfinal
    End If
    
    ' Marca en hoja Samples que se gener� calibraci�n
    Sheets("Samples").Range("AA32").value = "SI"
    
    ' --- VERIFICACI�N ARCHIVO EXISTENTE ---
    ' Comprueba si ya existe un PDF para este lote
    If Dir(rutacalibrarfinal & "\" & namepdf) <> "" Then
        Dim respuesta As VbMsgBoxResult
        ' Pide confirmaci�n para reemplazar
        respuesta = MsgBox("El archivo '" & namepdf & "' ya existe. �Deseas reemplazarlo?", vbYesNo + vbQuestion, "Confirmar Reemplazo")
        If respuesta = vbNo Then
            MsgBox "El archivo no se ha reemplazado.", vbInformation
            GoTo Line1  ' Sale si usuario elige no reemplazar
        End If
    End If
    
    ' --- EXPORTACI�N A PDF ---
    ' Selecciona hoja Criterios y exporta como PDF
    Sheets("Criterios").Activate
    ChDir rutacalibrarfinal  ' Cambia directorio actual
    ActiveSheet.ExportAsFixedFormat _
        Type:=xlTypePDF, _
        Filename:=rutacalibrarfinal & "\" & namepdf, _
        Quality:=xlQualityStandard, _
        IncludeDocProperties:=False, _
        IgnorePrintAreas:=False, _
        OpenAfterPublish:=False  ' No abrir despu�s de guardar
        
    ' Desactiva c�lculos y muestra confirmaci�n
    Worksheets("Criterios").EnableCalculation = False
    MsgBox "La exportaci�n se ha hecho bien", vbInformation
    
Line1:
    ' --- FINALIZACI�N ---
    ' Vuelve a hoja Samples y desactiva c�lculos
    Worksheets("Samples").Activate
    Worksheets("Criterios").EnableCalculation = False
End Sub

