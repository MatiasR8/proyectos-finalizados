Attribute VB_Name = "ExportQC_Mat"
Sub exportarQC(n As Integer)

    ' --- INICIALIZACIÓN DE DATOS ---
    ' Obtiene ruta base desde hoja Samples
    rutaexportqc = Sheets("Samples").Range("rutaexportqc").value
    
    ' Obtiene metadatos desde hoja CCD
    analyst = Sheets("CCD").Range("analyst").value
    method = Split(Sheets("CCD").Range("metodo").value, "-")(0) & "-" & Split(Sheets("CCD").Range("metodo").value, "-")(1)
    equipo = Sheets("CCD").Range("equipo").value
    revision = Sheets("CCD").Range("revision").value
    
    ' Verifica si el batch contiene calibración
    If Sheets("CCD").Range("BC11").value <> "" Then
        calibration = "Yes"
    Else
        calibration = "No"
    End If
    
    ' --- VERIFICACIÓN DE DATOS QC ---
    ' Comprueba si hay datos QC en el rango especificado
    Set rng = ThisWorkbook.Sheets("CCD").Range("AE58:AM58")
    allempty = True
    For Each cell In rng
        If cell.value <> "" Then
            allempty = False
            Exit For
        End If
    Next cell
    
    ' Si no hay datos, termina la ejecución
    If allempty Then
        GoTo Line1
    End If

    ' --- PREPARACIÓN ARCHIVO EXPORTACIÓN ---
    ' Genera nombre de archivo con fecha y nombre de batch
    nombrearchive = Format(Date, "dd-mm-yyyy") & "_" & Replace(Replace(Split(Sheets("CCD").Range("batch").value, ".")(0), "(", "-"), ")", "") & "_" & n & ".csv"
    archivepath = rutaexportqc & "QC\"
    
    ' Desprotege hoja CCD para permitir operaciones
    Sheets("CCD").Unprotect Password:="0000"
    
    ' Verifica existencia de rutas
    If Dir(rutaexportqc, vbDirectory) = "" Then
        MsgBox "Hay un error en la ruta de exportación", vbinfo
        GoTo Line1
    End If
    
    ' Crea directorio si no existe
    If Dir(archivepath, vbDirectory) = "" Then
        MkDir archivepath
    End If

    ' --- CREACIÓN ARCHIVO CSV ---
    ' Crea nuevo libro y lo guarda como CSV
    Set WB = ActiveWorkbook
    Workbooks.Add
    Set WBarchive = ActiveWorkbook
    WBarchive.SaveAs Filename:=archivepath & "\" & nombrearchive, FileFormat:=xlCSV, CreateBackup:=False
    
    ' --- EXPORTACIÓN DE DATOS ---
    ' Copia parámetros del método (columnas E e I)
    WB.Activate
    Sheets("CCD").Select
    Range("E58:E208,I58:I208").Select
    Selection.Copy
    WBarchive.Activate
    Cells(1, 1).Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, Skipblanks _
        :=False, Transpose:=False
    
    ' Determina última fila con datos
    Range("A2").Select
    Selection.End(xlDown).Select
    lastRow = Selection.row
    
    ' Añade metadatos como columnas adicionales
    Cells(1, 3).value = "Analyst"
    Range(Cells(2, 3), Cells(lastRow, 3)).Select
    Selection.FormulaR1C1 = analyst
    
    Cells(1, 4).value = "Method"
    Range(Cells(2, 4), Cells(lastRow, 4)).Select
    Selection.FormulaR1C1 = method
    
    Cells(1, 5).value = "Equipment"
    Range(Cells(2, 5), Cells(lastRow, 5)).Select
    Selection.FormulaR1C1 = equipo
    
    Cells(1, 6).value = "Calibration"
    Range(Cells(2, 6), Cells(lastRow, 6)).Select
    Selection.FormulaR1C1 = calibration
    
    Cells(1, 7).value = "revision"
    Range(Cells(2, 7), Cells(lastRow, 7)).Select
    Selection.FormulaR1C1 = revision
    
    WBarchive.Save
    
    ' Manejo de errores para copia de resultados
    On Error GoTo Line2
        
    ' Copia resultados QC (solo celdas visibles)
    WB.Activate
    Sheets("CCD").Select
    Range("AE58:AK208").SpecialCells(xlCellTypeVisible).Select
    Selection.Copy
    WBarchive.Activate
    Cells(1, 8).Select
    WB.Activate
    Sheets("CCD").Select
    Range("AE58:Ak208").SpecialCells(xlCellTypeVisible).Select
    Selection.Copy
    WBarchive.Activate
    Cells(1, 8).Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, Skipblanks _
        :=False, Transpose:=False
    
    ' Mensaje de confirmación
    MsgBox "los QC se han exportado bien", vbOKOnly

Line2:
    ' Finalización - guarda y cierra archivo
    Cells(1, 1).Select
    WBarchive.Save
    WBarchive.Close
    Sheets("CCD").Protect Password:="0000"
    
Line1:
    ' Punto de salida sin operaciones
End Sub

Sub exportarCM(n As Integer)
    ' --- ESTRUCTURA SIMILAR A exportarQC ---
    ' La macro sigue el mismo patrón que exportarQC pero para datos CM
    
    ' Las principales diferencias son:
    ' 1. Usa el rango AM58:AZ58 para verificar datos existentes
    ' 2. Exporta datos del rango AM58:AZ208
    ' 3. Nombre de archivo sin sufijo numérico (_n)
    ' 4. Mensaje final específico para "Matriciales"
    
    ' El resto de la estructura (inicialización, creación archivo, etc.)
    ' es idéntica a la macro exportarQC

    rutaexportCM = Sheets("Samples").Range("rutaexportQC").value
    analyst = Sheets("CCD").Range("analyst").value
    method = Split(Sheets("CCD").Range("metodo").value, "-")(0) & "-" & Split(Sheets("CCD").Range("metodo").value, "-")(1)
    equipo = Sheets("CCD").Range("equipo").value
    If Sheets("CCD").Range("BC11").value <> "" Then
        calibration = "Yes"
    Else
        calibration = "No"
    End If
    
    Set rng = ThisWorkbook.Sheets("CCD").Range("AM58:AZ58")
    allempty = True
    For Each cell In rng
        If cell.value <> "" Then
            allempty = False
            Exit For
        End If
    Next cell
    If allempty Then
        GoTo Line1
    End If
    
    nombrearchive = Format(Date, "dd-mm-yyyy") & "_" & Replace(Replace(Split(Sheets("CCD").Range("batch").value, ".")(0), "(", "-"), ")", "") & ".csv"
    archivepath = rutaexportCM & "CM\"
    Sheets("CCD").Unprotect Password:="0000"
    
    If Dir(rutaexportCM, vbDirectory) = "" Then
        MsgBox "Hay un error en la ruta de exportación", vbinfo
        GoTo Line1
    End If
    
    If Dir(archivepath, vbDirectory) = "" Then
        MkDir archivepath
    End If
    
    Set WB = ActiveWorkbook
    Workbooks.Add
    Set WBarchive = ActiveWorkbook
    WBarchive.SaveAs Filename:=archivepath & "\" & nombrearchive, FileFormat:=xlCSV, CreateBackup:=False
    
    WB.Activate
    Sheets("CCD").Select
    Range("E58:E208,I58:I208").Select
    Selection.Copy
    WBarchive.Activate
    Cells(1, 1).Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, Skipblanks _
        :=False, Transpose:=False
    
    Range("A2").Select
    Selection.End(xlDown).Select
    lastRow = Selection.row
    

    Cells(1, 3).value = "Analyst"
    Range(Cells(2, 3), Cells(lastRow, 3)).Select
    Selection.FormulaR1C1 = analyst
    Cells(1, 4).value = "Method"
    Range(Cells(2, 4), Cells(lastRow, 4)).Select
    Selection.FormulaR1C1 = method
    Cells(1, 5).value = "Equipment"
    Range(Cells(2, 5), Cells(lastRow, 5)).Select
    Selection.FormulaR1C1 = equipo
    Cells(1, 6).value = "Calibration"
    Range(Cells(2, 6), Cells(lastRow, 6)).Select
    Selection.FormulaR1C1 = calibration
    Cells(1, 7).value = "revision"
    Range(Cells(2, 7), Cells(lastRow, 7)).Select
    Selection.FormulaR1C1 = revision
    WBarchive.Save
        
    On Error GoTo Line2
    
    WB.Activate
    Sheets("CCD").Select
    Range("AM58:AZ208").SpecialCells(xlCellTypeVisible).Select
    Selection.Copy
    WBarchive.Activate
    Cells(1, 8).Select
    WB.Activate
    Sheets("CCD").Select
    Range("AM58:AZ208").SpecialCells(xlCellTypeVisible).Select
    Selection.Copy
    WBarchive.Activate
    Cells(1, 8).Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, Skipblanks _
        :=False, Transpose:=False
        
   MsgBox "los Matriciales se han exportado bien", vbOKOnly
   
Line2:
    Cells(1, 1).Select
    WBarchive.Save
    WBarchive.Close
    Sheets("CCD").Protect Password:="0000"
    
Line1:

End Sub





