Attribute VB_Name = "Exportsample"
Sub Exportar()
' Macro para exportar la lista de muestras de la hoja "Sample" y los controles QC/CM

    ' 1. COPIA DE CARPETA BASE
    ' Ejecuta la macro que copia la carpeta con los datos primarios
    Call CopiarCarpeta2
    
    ' 2. PROCESAMIENTO ESPECIAL PARA PLAGUICIDAS
    ' Si el método es CGM/031 o CGM/019, ejecuta rutina especial
    If Left(Sheets("CCD").Range("J12").value, 7) = "CGM/031" Or Left(Sheets("CCD").Range("J12").value, 7) = "CGM/019" Then
        Call Plaguicidas
    End If
    
    ' Optimización: desactiva actualización de pantalla
    Application.ScreenUpdating = False

    ' 3. PREPARACIÓN ARCHIVO DE EXPORTACIÓN
    ' Obtiene ruta y nombre base del archivo
    ruta = Worksheets("Samples").Range("rutaexport")
    nombre = Split(Sheets("CCD").Range("batch").value, ".")(0)
    
    ' Desprotege hoja de exportación
    Sheets("Exportacion").Unprotect Password:="0000"
    
    ' 4. VALIDACIÓN DE DATOS
    ' Verifica que haya muestras para exportar
    If Sheets("Exportacion").Range("A1").value = "" Then
        MsgBox "No hay muestras a exportar", vbinfo
        GoTo Line1  ' Salta al final si no hay datos
    End If
    
    ' 5. GENERACIÓN DE NOMBRE ÚNICO
    ' Busca el próximo número disponible para evitar sobrescribir
    n = 0
    Do While Dir(ruta & nombre & "_" & n & ".txt") <> ""
        n = n + 1
    Loop
    
    ' 6. PREPARACIÓN DE DATOS
    ' Crea libro temporal y copia hoja de exportación
    Set shtAExportar = ThisWorkbook.Worksheets("Exportacion")
    Set wbkTemporal = Application.Workbooks.Add
    shtAExportar.Copy Before:=wbkTemporal.Worksheets(wbkTemporal.Worksheets.Count)
    
    ' 7. LIMPIEZA DE DATOS
    ' Elimina filas vacías del archivo a exportar
    If Sheets("Exportacion").Range("A2").value <> "" Then
        Range("A1").End(xlDown).Select
    Else
        Range("A1").Select
    End If
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Offset(1, 0).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Delete Shift:=xlUp
    
    ' 8. EXPORTACIÓN PRINCIPAL
    ' Guarda en dos ubicaciones: ruta principal y C:\temp\report\
    Application.DisplayAlerts = False
    wbkTemporal.SaveAs Filename:=ruta & nombre & "_" & n & ".txt", FileFormat:=xlText
    wbkTemporal.SaveAs Filename:="C:\temp\report\" & nombre & "_" & n & ".txt", FileFormat:=xlText
    nombre = ruta & nombre & ".txt"
    Application.DisplayAlerts = True
        
    ' 9. LIMPIEZA FINAL
    ' Cierra archivo temporal y limpia hoja de exportación
    wbkTemporal.Close SaveChanges:=False
    Sheets("Exportacion").Activate
    Range("A1:J1").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.ClearContents
    Sheets("Exportacion").Protect Password:="0000"
    
    ' 10. EXPORTACIÓN DE CONTROLES
    ' Ejecuta macros para exportar QC y CM
    Worksheets("Samples").Activate
    exportarQC n  ' Exporta controles de calidad
    exportarCM n  ' Exporta controles matriciales
    
Line1:
    ' Finalización: reactiva actualización y vuelve a hoja Samples
    Worksheets("Samples").Activate
    Application.ScreenUpdating = True
End Sub
