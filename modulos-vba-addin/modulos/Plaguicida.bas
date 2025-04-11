Attribute VB_Name = "Plaguicida"
Sub Plaguicidas()
' Macro para identificar y alertar sobre muestras con parámetros de plaguicidas y HPA
' que no han sido subidos correctamente al sistema

    ' --- CONFIGURACIÓN INICIAL ---
    ' Establece referencias a las hojas de trabajo
    Set wsExp = ThisWorkbook.Sheets("Exportacion")       ' Hoja con datos exportados
    Set wsParam = ThisWorkbook.Sheets("Parámetros_Barrido")  ' Hoja de parámetros
    
    ' --- PREPARACIÓN DE DATOS ---
    ' Encuentra la última fila con datos en columna D de Exportacion
    lastRowExp = wsExp.Cells(wsExp.Rows.Count, "D").End(xlUp).row
    
    ' --- IDENTIFICACIÓN DE PLAGUICIDAS ---
    ' Crea diccionario para almacenar códigos de muestra únicos
    Set dict = CreateObject("Scripting.Dictionary")
    
    ' Variables para control y mensaje
    encontrado = False
    mensaje = "Estas muestras tienen parámetros que no se han subido:" & vbNewLine
    
    ' Recorre todas las filas buscando "Plaguicidas" o "Plaguicidas totales"
    For i = 1 To lastRowExp
        ' Verifica si la celda en columna D contiene los valores buscados
        If wsExp.Cells(i, 4).value = "Plaguicidas" Or wsExp.Cells(i, 4).value = "Plaguicidas totales" Then
            ' Agrega el código de muestra (columna B) al diccionario si no existe
            If Not dict.Exists(wsExp.Cells(i, 2).value) Then
                dict.Add wsExp.Cells(i, 2).value, Nothing
                encontrado = True
                ' Construye mensaje con códigos de muestra
                mensaje = mensaje & "- " & wsExp.Cells(i, 2).value & vbNewLine
            End If
        End If
    Next i
    
    ' --- IDENTIFICACIÓN DE HPA (Hidrocarburos Policíclicos Aromáticos) ---
    ' Crea segundo diccionario para HPA
    Set dict2 = CreateObject("Scripting.Dictionary")
    encontrado2 = False
    
    ' Recorre filas buscando "HPA" y excluyendo "Agua de consumo"
    For i = 1 To lastRowExp
        If wsExp.Cells(i, 4).value = "HPA" Then
            ' Verifica que no sea agua de consumo y agrega al diccionario
            If Not dict2.Exists(wsExp.Cells(i, 2).value) And wsExp.Cells(i, 8).value <> "Agua de consumo" Then
                dict2.Add wsExp.Cells(i, 2).value, Nothing
                encontrado2 = True
                ' Amplía el mensaje con estos códigos
                mensaje = mensaje & "- " & wsExp.Cells(i, 2).value & vbNewLine
            End If
        End If
    Next i
    
    ' --- MOSTRAR RESULTADOS ---
    ' Si encontró muestras con plaguicidas no subidos, muestra alerta
    If encontrado Or encontrado2 Then
        MsgBox mensaje, vbExclamation, "Alerta de Plaguicidas"
    End If
End Sub
