Attribute VB_Name = "ReportMuestras"
Sub ExportReport()
' Macro para exportar reportes a PDF con manejo de versiones y rutas espec�ficas

    ' --- CONFIGURACI�N INICIAL ---
    Set activeWB = ActiveWorkbook  ' Guarda referencia al libro activo
    Worksheets("Report").EnableCalculation = True  ' Asegura c�lculos actualizados

    ' --- PREPARACI�N DE NOMBRES Y RUTAS ---
    ' Genera nombre base para archivos (reemplaza caracteres problem�ticos)
    archivename = "R" & Replace(Replace(Split(Sheets("CCD").Range("batch").value, ".")(0), "(", "-"), ")", "")
    
    ' Obtiene ruta base para exportaci�n desde hoja Samples
    rutaexport = Sheets("Samples").Range("rutaexportreport").value
    
    ' Determina ruta final seg�n tipo de m�todo
    If Sheets("CCD").Range("J12").value <> "CGM/019-pcbbde" And Sheets("CCD").Range("J12").value <> "CGM/031-a-CP" Then
        ' Ruta est�ndar: usa nombre de batch
        rutaexportfinal = rutaexport & Replace(Replace(Split(Sheets("CCD").Range("batch").value, ".")(0), "(", "-"), ")", "")
    Else
        ' Ruta especial para m�todos espec�ficos: usa campo E13
        rutaexportfinal = rutaexport & Replace(Replace(Split(Sheets("CCD").Range("E13").value, ".")(0), "(", "-"), ")", "")
    End If
    
    ' Determina nombre del PDF seg�n tipo de m�todo
    If Sheets("CCD").Range("J12").value = "CGM/040-a" Or Sheets("CCD").Range("J12").value = "CGM/041-a" Or Sheets("CCD").Range("J12").value = "CGM/026-a" Then
        ' Nombre especial para m�todos espec�ficos
        PDFname = Sheets("Samples").Range("E6").value & "." & Replace(Sheets("Samples").Range("samplename").value, "/", ".") & ".pdf"
    Else
        ' Nombre est�ndar para otros m�todos
        PDFname = Sheets("Samples").Range("SampleIDs").value & "." & Replace(Sheets("Samples").Range("samplename").value, "/", ".") & ".pdf"
    End If

    ' --- MANEJO DE VERSIONES EXISTENTES ---
    Application.DisplayAlerts = False  ' Evita alertas durante el proceso
    
    ' Verifica si el archivo ya existe
    If Dir(rutaexportfinal & "\" & PDFname, vbDirectory) <> "" Then
        ' Opciones para manejar archivos existentes
        result = MsgBox("Esta muestra ya se ha guardado, �Quieres crear un nuevo PDF?" & vbCrLf & _
               "    �Si: Crea un nuevo PDF con otro nombre." & vbCrLf & _
               "    �No: Sobrescribe el PDF actual.", vbYesNoCancel)
        
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

    ' --- VERIFICACI�N/CREACI�N DE DIRECTORIOS ---
    ' Comprueba existencia de directorio final
    If Dir(rutaexportfinal, vbDirectory) = "" Then
        ' Verifica primero la ruta base
        If Dir(rutaexport, vbDirectory) = "" Then
            MsgBox "Hay un error en la ruta de exportaci�n", vbInformation
            GoTo Line1  ' Sale si la ruta base no existe
        End If
        ' Crea directorio espec�fico para este lote
        MkDir (rutaexportfinal)
    End If

    ' --- EXPORTACI�N A PDF ---
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
        OpenAfterPublish:=False  ' No abre el PDF despu�s de guardar
    
    Application.DisplayAlerts = True  ' Restaura alertas

Line1:
    ' --- FINALIZACI�N ---
    ' Vuelve al estado original
    activeWB.Activate
    Worksheets("Samples").Activate
    Worksheets("Report").EnableCalculation = False
End Sub

