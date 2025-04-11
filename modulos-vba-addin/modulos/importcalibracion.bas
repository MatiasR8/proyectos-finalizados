Attribute VB_Name = "importcalibracion"
Sub Importar_calibracion()
' Macro para importar criterios de calibración desde un archivo externo
' y filtrarlos según el método seleccionado

    ' --- CONFIGURACIÓN INICIAL ---
    ' Desactiva actualización de pantalla para mejor rendimiento
    Application.ScreenUpdating = False

    ' Obtiene nombre del libro activo y prepara hoja Criterios
    nombre_excel = ActiveWorkbook.Name
    Sheets("Criterios").Unprotect Password:="0000"
    Worksheets("Criterios").EnableCalculation = True
    
    ' Limpia área de importación previa
    Range("D20:F26").ClearContents
    
    ' --- PREPARACIÓN DE RUTA Y FILTRO ---
    ' Construye ruta completa del archivo de criterios
    ruta_archivo = Sheets("Criterios").Range("rutacriterios").value & "Tabla_Criterios"
    
    ' Obtiene método de calibración a filtrar
    metodo = Sheets("Criterios").Range("Calibracion").value
    
    ' --- IMPORTACIÓN DEL ARCHIVO ---
    ' Abre archivo de criterios como texto (formato específico)
    Workbooks.OpenText Filename:= _
        ruta_archivo & ".xlsx" _
        , Origin:=1252, StartRow:=1, DataType:=xlDelimited, TextQualifier:= _
        xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, Semicolon:=False, _
        Comma:=False, Space:=False, Other:=False, FieldInfo:=Array(Array(1, 1), _
        Array(2, 1), Array(3, 1), Array(4, 1), Array(5, 1), Array(6, 1), Array(7, 1), Array(8, 1), _
        Array(9, 1), Array(10, 1), Array(13, 1), Array(14, 1)), TrailingMinusNumbers:=True
    
    ' --- PROCESAMIENTO DE DATOS ---
    ' Selecciona y filtra datos en el archivo abierto
    Windows("Tabla_Criterios.xlsx").Activate
    Sheets("Sheet1").Select
    
    ' Selecciona rango de datos (desde B2:D2 hasta final de datos)
    Range("b2:D2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Set importrange = Selection
    
    ' Aplica filtro por método de calibración (columna E)
    Range("E1").Select
    ActiveSheet.Range("Table1").AutoFilter Field:=5, Criteria1:=metodo
    
    ' Copia solo las filas visibles (filtradas)
    importrange.SpecialCells(xlCellTypeVisible).Select
    Selection.Copy
    
    ' --- PEGADO DE DATOS ---
    ' Vuelve al libro original y pega los valores
    Windows(nombre_excel).Activate
    Sheets("Criterios").Range("d20").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, Skipblanks _
        :=False, Transpose:=False
    
    ' --- LIMPIEZA FINAL ---
    ' Cierra archivo de criterios sin guardar
    Application.CutCopyMode = False
    Windows("Tabla_Criterios.xlsx").Activate
    Range("A1").Select
    ActiveWindow.Close False
    
    ' Protege hoja y selecciona celda inicial
    Sheets("Criterios").Protect Password:="0000"
    Sheets("Criterios").Activate
    Range("D20").Select
    
    ' Reactiva actualización de pantalla
    Application.ScreenUpdating = True
End Sub




