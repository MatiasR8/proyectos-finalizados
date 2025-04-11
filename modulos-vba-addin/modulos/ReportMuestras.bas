Attribute VB_Name = "ReportMuestras"
Sub ExportReport()
' Macro para exportar reportes a PDF con manejo de versiones y rutas específicas

    ' --- CONFIGURACIÓN INICIAL ---
    Set activeWB = ActiveWorkbook  ' Guarda referencia al libro activo
    Worksheets("Report").EnableCalculation = True  ' Asegura cálculos actualizados

    ' --- PREPARACIÓN DE NOMBRES Y RUTAS ---
    ' Genera nombre base para archivos (reemplaza caracteres problemáticos)
    archivename = "R" & Replace(Replace(Split(Sheets("CCD").Range("batch").value, ".")(0), "(", "-"), ")", "")
    
    ' Obtiene ruta base para exportación desde hoja Samples
    rutaexport = Sheets("Samples").Range("rutaexportreport").value
    
    ' Determina ruta final según tipo de método
    If Sheets("CCD").Range("J12").value <> "CGM/019-pcbbde" And Sheets("CCD").Range("J12").value <> "CGM/031-a-CP" Then
        ' Ruta estándar: usa nombre de batch
        rutaexportfinal = rutaexport & Replace(Replace(Split(Sheets("CCD").Range("batch").value, ".")(0), "(", "-"), ")", "")
    Else
        ' Ruta especial para métodos específicos: usa campo E13
        rutaexportfinal = rutaexport & Replace(Replace(Split(Sheets("CCD").Range("E13").value, ".")(0), "(", "-"), ")", "")
    End If
    
    ' Determina nombre del PDF según tipo de método
    If Sheets("CCD").Range("J12").value = "CGM/040-a" Or Sheets("CCD").Range("J12").value = "CGM/041-a" Or Sheets("CCD").Range("J12").value = "CGM/026-a" Then
        ' Nombre especial para métodos específicos
        PDFname = Sheets("Samples").Range("E6").value & "." & Replace(Sheets("Samples").Range("samplename").value, "/", ".") & ".pdf"
    Else
        ' Nombre estándar para otros métodos
        PDFname = Sheets("Samples").Range("SampleIDs").value & "." & Replace(Sheets("Samples").Range("samplename").value, "/", ".") & ".pdf"
    End If

    ' --- MANEJO DE VERSIONES EXISTENTES ---
    Application.DisplayAlerts = False  ' Evita alertas durante el proceso
    
    ' Verifica si el archivo ya existe
    If Dir(rutaexportfinal & "\" & PDFname, vbDirectory) <> "" Then
        ' Opciones para manejar archivos existentes
        result = MsgBox("Esta muestra ya se ha guardado, ¿Quieres crear un nuevo PDF?" & vbCrLf & _
               "    ·Si: Crea un nuevo PDF con otro nombre." & vbCrLf & _
               "    ·No: Sobrescribe el PDF actual.", vbYesNoCancel)
        
        If result = vbYes Then
            ' Genera nombres incrementales (ej: _2, _3)
            contador = 2
            Do While Dir(rutaexportfinal & "\" & PDFname, vbDirectory) <> ""
                PDFname = Sheets("Samples").Range("SampleIDs").value & "." & _
                         Replace(Sheets("Samples").Range("samplename").value, "/", ".") & _
                         "_" & contador & ".pdf"
                contador = contador + 1
            Loop
        ElseIf result = vbNo Then
            ' Mantiene nombre original para sobrescribir
            If Sheets("CCD").Range("J12").value = "CGM/040-a" Or _
               Sheets("CCD").Range("J12").value = "CGM/041-a" Or _
               Sheets("CCD").Range("J12").value = "CGM/026-a" Then
                PDFname = Sheets("Samples").Range("E6").value & "." & _
                         Replace(Sheets("Samples").Range("samplename").value, "/", ".") & ".pdf"
            Else
                PDFname = Sheets("Samples").Range("SampleIDs").value & "." & _
                         Replace(Sheets("Samples").Range("samplename").value, "/", ".") & ".pdf"
            End If
        ElseIf result = vbCancel Then
            ' Sale de la macro si usuario cancela
            Application.DisplayAlerts = True
            Exit Sub
        End If
    End If

    ' --- VERIFICACIÓN/CREACIÓN DE DIRECTORIOS ---
    ' Comprueba existencia de directorio final
    If Dir(rutaexportfinal, vbDirectory) = "" Then
        ' Verifica primero la ruta base
        If Dir(rutaexport, vbDirectory) = "" Then
            MsgBox "Hay un error en la ruta de exportación", vbInformation
            GoTo Line1  ' Sale si la ruta base no existe
        End If
        ' Crea directorio específico para este lote
        MkDir (rutaexportfinal)
    End If

    ' --- EXPORTACIÓN A PDF ---
    activeWB.Activate
    Sheets("Report").Activate
    ChDir rutaexportfinal  ' Cambia directorio actual
    
    ' Exporta hoja Report como PDF
    ActiveSheet.ExportAsFixedFormat _
        Type:=xlTypePDF, _
        Filename:=rutaexportfinal & "\" & PDFname, _
        Quality:=xlQualityStandard, _
        IncludeDocProperties:=False, _
        IgnorePrintAreas:=False, _
        OpenAfterPublish:=False  ' No abre el PDF después de guardar
    
    Application.DisplayAlerts = True  ' Restaura alertas

Line1:
    ' --- FINALIZACIÓN ---
    ' Vuelve al estado original
    activeWB.Activate
    Worksheets("Samples").Activate
    Worksheets("Report").EnableCalculation = False
End Sub

