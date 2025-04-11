Attribute VB_Name = "CopiarCarpeta"
' Macro para copiar carpetas de resultados analíticos a ubicaciones de red específicas
' según diferentes criterios de métodos y configuraciones
Sub CopiarCarpeta2()
    
    ' --- Validación inicial ---
    ' Verifica si el contenido de la celda J9 es una fecha válida
    If Not IsDate(Sheets("CCD").Range("J9").value) Then
        ' Si no es fecha válida, muestra mensaje y sale de la subrutina
        MsgBox "El nombre del batch no tiene el formato adecuado, copia la carpeta a red a mano.", vbCritical, "Error de formato"
        Exit Sub
    End If
    
    ' --- Lógica para métodos CG-MS específicos ---
    ' Comprueba si el método es CGM/040-a, CGM/041-a o CGM/026-a
    If Sheets("CCD").Range("J12").value = "CGM/040-a" Or Sheets("CCD").Range("J12").value = "CGM/041-a" Or Sheets("CCD").Range("J12").value = "CGM/026-a" Then
        ' Pregunta al usuario si quiere guardar con fecha actual
        respuesta = MsgBox("¿Quieres guardar el PDF con el día actual?", vbYesNoCancel)
        
        ' Opción Sí: copia toda la carpeta
        If respuesta = vbYes Then
            ' Construye ruta de origen reemplazando caracteres problemáticos
            origen = Sheets("Samples").Range("rutaexportreport").value & _
                 Replace(Replace(Split(Sheets("CCD").Range("batch").value, ".")(0), "(", "-"), ")", "")
        
            ' Construye ruta de destino usando valores de varias celdas
            destino = "\\ruta-censurada\8. LABORATORIO CROMATOGRAFIA CS\4. DATOS PRIMARIOS\RESULTADOS 20" & _
                  Sheets("CCD").Range("H9").value & "\Métodos CG-MS\" & _
                  Sheets("CCD").Range("H11").value & "\" & _
                  Sheets("CCD").Range("H8").value & "." & _
                  Sheets("CCD").Range("H10").value & "\"
        
            ' Crea objeto para manejo de archivos
            Set FSO = CreateObject("Scripting.FileSystemObject")
            
            ' Verifica si existe la carpeta destino, si no existe la crea
            If Not FSO.FolderExists(destino) Then
                FSO.CreateFolder destino
            End If
        
            ' Copia toda la carpeta de origen a destino
            FSO.CopyFolder origen, destino
       
        ' Opción No: copia archivos individualmente con subcarpeta adicional
        ElseIf respuesta = vbNo Then
            ' Verifica si hay valor en H13 (para cambio de mes)
            If Sheets("CCD").Range("H13").value <> 0 Then
                origen = Sheets("Samples").Range("rutaexportreport").value & _
                     Replace(Replace(Split(Sheets("CCD").Range("batch").value, ".")(0), "(", "-"), ")", "")
        
                ' Construye ruta destino incluyendo el valor de H18 como subcarpeta
                destino = "\\ruta-censurada\8. LABORATORIO CROMATOGRAFIA CS\4. DATOS PRIMARIOS\RESULTADOS 20" & _
                      Sheets("CCD").Range("H9").value & "\Métodos CG-MS\" & _
                      Sheets("CCD").Range("H11").value & "\" & _
                      Sheets("CCD").Range("H8").value & "." & _
                      Sheets("CCD").Range("H10").value & "\" & _
                      Sheets("CCD").Range("H18").value
        
                ' Crea objeto para manejo de archivos
                Set FSO = CreateObject("Scripting.FileSystemObject")
                
                ' Verifica si existe la carpeta destino, si no existe la crea
                If Not FSO.FolderExists(destino) Then
                    FSO.CreateFolder destino
                End If
        
                ' Copia archivos individualmente (no la carpeta completa)
                For Each archivo In FSO.GetFolder(origen).Files
                    FSO.CopyFile archivo.Path, destino & "\" & archivo.Name
                Next archivo
            
            ' Si H13 es 0 (cambio de mes), indica que se debe hacer manualmente
            ElseIf Sheets("CCD").Range("H13").value = 0 Then
                MsgBox "Cambio de mes, hacer a mano", vbOKOnly
            End If
        
        ' Opción Cancelar: sale de la subrutina
        ElseIf respuesta = vbCancel Then
            Application.DisplayAlerts = True
            Exit Sub
        End If
    
    ' --- Lógica para métodos CG estándar ---
    ' Comprueba si el método es CG/025-a o CG/026-a
    ElseIf Sheets("CCD").Range("J12").value = "CG/025-a" Or Sheets("CCD").Range("J12").value = "CG/026-a" Then
        ' Construye ruta de origen
        origen = Sheets("Samples").Range("rutaexportreport").value & _
                    Replace(Replace(Split(Sheets("CCD").Range("batch").value, ".")(0), "(", "-"), ")", "")
        
        ' Construye ruta destino específica para métodos CG
        destino = "\\ruta-censurada\8. LABORATORIO CROMATOGRAFIA CS\4. DATOS PRIMARIOS\RESULTADOS 20" & _
                    Sheets("CCD").Range("H9").value & "\Método CG\" & _
                    Sheets("CCD").Range("H11").value & "\" & _
                    Sheets("CCD").Range("H8").value & "." & _
                    Sheets("CCD").Range("H10").value & "\"
        
        ' Crea objeto para manejo de archivos
        Set FSO = CreateObject("Scripting.FileSystemObject")
            
        ' Verifica si existe la carpeta destino, si no existe la crea
        If Not FSO.FolderExists(destino) Then
            FSO.CreateFolder destino
        End If
        
        ' Verifica si existe la carpeta origen antes de copiar
        If Not FSO.FolderExists(origen) Then
            MsgBox "La carpeta de origen no está creada, guarda primero alguna muestra"
            Exit Sub
        Else
            ' Copia toda la carpeta de origen a destino
            FSO.CopyFolder origen, destino
        End If
    
    ' --- Lógica para todos los demás métodos ---
    Else
        ' Construye ruta de origen (similar a casos anteriores)
        origen = Sheets("Samples").Range("rutaexportreport").value & _
                 Replace(Replace(Split(Sheets("CCD").Range("batch").value, ".")(0), "(", "-"), ")", "")
        
        ' Construye ruta destino para métodos no específicos
        destino = "\\ruta-censurada\8. LABORATORIO CROMATOGRAFIA CS\4. DATOS PRIMARIOS\RESULTADOS 20" & _
                  Sheets("CCD").Range("H9").value & "\Métodos CG-MS\" & _
                  Sheets("CCD").Range("H11").value & "\" & _
                  Sheets("CCD").Range("H8").value & "." & _
                  Sheets("CCD").Range("H10").value & "\"
        
        ' Crea objeto para manejo de archivos
        Set FSO = CreateObject("Scripting.FileSystemObject")
            
        ' Verifica si existe la carpeta destino, si no existe la crea
        If Not FSO.FolderExists(destino) Then
            FSO.CreateFolder destino
        End If
        
        ' Verifica si existe la carpeta origen antes de copiar
        If Not FSO.FolderExists(origen) Then
            MsgBox "La carpeta de origen no está creada, guarda primero alguna muestra"
            Exit Sub
        Else
            ' Copia toda la carpeta de origen a destino
            FSO.CopyFolder origen, destino
        End If
    End If

End Sub

' Subrutina para verificar existencia de archivo PDF de criterios
Sub ComprobarCriterios()
    ' Genera nombre del archivo PDF concatenando "Criterios_" con el nombre del batch
    Dim namepdf As String
    namepdf = "Criterios_" & Split(Sheets("CCD").Range("batch").value, ".")(0) & ".pdf"
    
    ' Obtiene ruta base desde hoja Criterios
    Dim ruta As String
    ruta = Sheets("Criterios").Range("rutacalibrar").value
    
    ' Construye ruta completa reemplazando caracteres problemáticos
    Dim rutacalibrarfinal As String
    rutacalibrarfinal = ruta & Replace(Replace(Split(Sheets("CCD").Range("batch").value, ".")(0), "(", "-"), ")", "")
   
    ' Verifica si existe el archivo PDF en la ruta construida
    If Dir(rutacalibrarfinal & "\" & namepdf) <> "" Then
        ' Si existe, marca "SI" en la celda AA32 de la hoja Samples
        Sheets("Samples").Range("AA32").value = "SI"
    Else
        ' Si no existe, marca "NO" en la celda AA32 de la hoja Samples
        Sheets("Samples").Range("AA32").value = "NO"
    End If
End Sub
