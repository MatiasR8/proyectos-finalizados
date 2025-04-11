Attribute VB_Name = "ImportLims"
Sub Importar_archivo_lims()
' Macro para importar datos desde un archivo LIMS al libro actual
' Realiza m�ltiples operaciones de preparaci�n, importaci�n y procesamiento

    ' --- CONFIGURACI�N INICIAL ---
    ' Optimiza rendimiento desactivando actualizaciones
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    ' --- LIMPIEZA DE DATOS DE BARRIDO PREVIA ---
    ' Advertencia y limpieza de par�metros de barrido si existen
    If Sheets("Par�metros_Barrido").Range("A1") <> "" Then
        result = MsgBox("Los par�metros barridos guardados se borrar�n con la importaci�n, �Deseas continuar?", vbOKCancel)
        If result = vbCancel Then
                GoTo Line1  ' Salta al final si usuario cancela
        ElseIf result = vbOK Then
            ' Limpia hoja de par�metros de barrido
            Set wsLimpiar = Sheets("Par�metros_Barrido")
            wsLimpiar.Unprotect Password:="0000"
            wsLimpiar.Select
            Range("A1:N1").Select
            Range(Selection, Selection.End(xlDown)).Select
            Set limpieza = Selection
            limpieza.ClearContents
            wsLimpiar.Protect Password:="0000"
        End If
    End If
    
    ' --- IMPORTACIONES PREVIAS ---
    ' Importa datos de blancos
    Call Blancos
    
    ' Para m�todos espec�ficos, importa datos de controles C5-C40
    If Sheets("CCD").Range("J12").value = "CGM/040-a" Or Sheets("CCD").Range("J12").value = "CGM/041-a" Then
        Call c5c40
    End If
    
    ' Importa matriz de par�metros
    Call Importar_Matrix_parametros_MH_lims
    
    ' --- PREPARACI�N HOJA LIMS ---
    ' Guarda contexto actual y prepara hoja LIMS
    nombre_excel = ActiveWorkbook.Name
    activehoja = ActiveSheet.Name
    Sheets("LIMS").Unprotect Password:="0000"
    Worksheets("LIMS").EnableCalculation = True
    Worksheets("Overview").EnableCalculation = True
    
    ' Elimina tabla existente y limpia datos
    Sheets("LIMS").Select
    Set tbllims = ActiveSheet.ListObjects("limsimport")
    tbllims.Unlist
    Sheets("LIMS").Columns("B:M").ClearContents

    ' --- CONFIGURACI�N IMPORTACI�N ---
    ' Obtiene par�metros desde hoja Samples
    ruta_archivo = Sheets("Samples").Range("rutalims").value
    nombre_archivo = Sheets("Samples").Range("nombrearchivolims").value
    metodo = Split(Sheets("CCD").Range("metodo").value, "-")(0) & "*"  ' Formato para filtro
    
    ' --- IMPORTACI�N ARCHIVO LIMS ---
    ' Abre archivo LIMS como texto con configuraci�n espec�fica
    Workbooks.OpenText Filename:= _
        ruta_archivo & "\" & nombre_archivo _
        , Origin:=1252, StartRow:=1, DataType:=xlDelimited, TextQualifier:= _
        xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, Semicolon:=False, _
        Comma:=False, Space:=False, Other:=False, FieldInfo:=Array(Array(1, 1), _
        Array(2, 1), Array(3, 1), Array(4, 1), Array(5, 1), Array(6, 1), Array(7, 1), Array(8, 1), _
        Array(9, 1), Array(10, 1), Array(13, 1), Array(14, 1)), TrailingMinusNumbers:=True
    
    ' --- PREPARACI�N DATOS IMPORTADOS ---
    Windows(nombre_archivo).Activate
    ' A�ade fila de encabezados
    Rows("1:1").Insert Shift:=xlDown
    Range("A1").value = "A�o"
    Range("B1").value = "Muestra"
    Range("C1").value = "Cod.Param"
    Range("D1").value = "Parametro"
    Range("E1").value = "Cod.Met"
    Range("F1").value = "Metodo"
    Range("G1").value = "Cod.Matr"
    Range("H1").value = "Matriz"
    Range("I1").value = "Resultado"
    Range("J1").value = "Unidades"
    Range("K1").value = "F1"
    Range("L1").value = "F2"
    Range("M1").value = "LOQ"
    Range("N1").value = "Unidades.LOQ"
    
    ' Aplica filtro por m�todo
    Columns("A:N").AutoFilter
    ActiveSheet.Range("$A$1:$N$1000000").AutoFilter Field:=6, Criteria1:=metodo
    
    ' --- COPIA DATOS FILTRADOS ---
    On Error Resume Next
    Range("A:J,M:N").Select  ' Selecciona columnas relevantes (excluyendo K y L)
    Range("A:J;M:N").Select  ' Formato alternativo para selecci�n
    On Error GoTo 0
    Selection.Copy
    
    ' Pega datos en hoja LIMS
    Windows(nombre_excel).Activate
    Sheets("LIMS").Range("B1").PasteSpecial Paste:=xlPasteValues
    
    ' --- CONVERTIR DATOS EN TABLA ---
    Range("A1:M1").Select
    Range(Selection, Selection.End(xlDown)).Select
    ActiveSheet.ListObjects.Add(xlSrcRange, Selection, , xlYes).Name = "limsimport"
    
    ' --- LIMPIEZA FINAL ---
    Application.CutCopyMode = False
    Windows(nombre_archivo).Activate
    Range("A1").Select
    Selection.AutoFilter
    ActiveWindow.Close False  ' Cierra archivo LIMS sin guardar
    
    ' Reactiva c�lculos y protege hoja
    Worksheets("Overview").EnableCalculation = True
    Worksheets("LIMS").EnableCalculation = False
    Sheets("LIMS").Protect Password:="0000"
    
    ' Restaura configuraci�n Excel
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
        
    ' --- PROCESAMIENTOS POST-IMPORTACI�N ---
    Sheets(activehoja).Activate  ' Vuelve a hoja original
    
    ' Ejecuta validaciones adicionales
    Call ComprobarCriterios  ' Identifica gemelas
    Call calculate            ' Comprueba criterios exportados para m�todos 42/43

Line1:
    ' Punto de salida alternativo
    Sheets("CCD").Select
    Application.ScreenUpdating = False
End Sub


