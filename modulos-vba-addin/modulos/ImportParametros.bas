Attribute VB_Name = "ImportParametros"
Sub Importar_Matrix_parametros_MH_lims()
' Macro para importar la matriz de par�metros MH-LIMS desde un archivo externo
' y filtrarla seg�n el m�todo especificado en la hoja CCD

    ' --- CONFIGURACI�N INICIAL ---
    ' Desactiva actualizaci�n de pantalla para mejor rendimiento
    Application.ScreenUpdating = False
    
    ' --- PREPARACI�N DEL ENTORNO ---
    ' Guarda nombre del libro actual y desprotege hoja InformeFinal
    nombre_excel = ActiveWorkbook.Name
    Sheets("InformeFinal").Unprotect Password:="0000"
    Sheets("InformeFinal").Select
    
    ' Obtiene referencia a la tabla existente (matrixmhlims)
    Set tbllims = ActiveSheet.ListObjects("matrixmhlims")
    
    ' Limpia el �rea de importaci�n previa (columnas Q a W)
    Sheets("InformeFinal").Range("Q4:W200").ClearContents

    ' --- CONFIGURACI�N DE IMPORTACI�N ---
    ' Obtiene ruta del archivo y m�todo desde las hojas
    ruta_archivo = Sheets("Samples").Range("rutaparametros").value
    metodo = Sheets("CCD").Range("metodo").value
    
    ' --- IMPORTACI�N DEL ARCHIVO ---
    ' Abre el archivo de par�metros como texto con configuraci�n espec�fica
    Workbooks.OpenText Filename:= _
        ruta_archivo & ".xlsx" _
        , Origin:=1252, StartRow:=1, DataType:=xlDelimited, TextQualifier:= _
        xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, Semicolon:=False, _
        Comma:=False, Space:=False, Other:=False, FieldInfo:=Array(Array(1, 1), _
        Array(2, 1), Array(3, 1), Array(4, 1), Array(5, 1), Array(6, 1), Array(7, 1), Array(8, 1), _
        Array(9, 1), Array(10, 1), Array(13, 1), Array(14, 1)), TrailingMinusNumbers:=True
    
    ' --- PROCESAMIENTO DE DATOS ---
    ' Activa el archivo reci�n abierto (nombre hardcodeado - posible mejora)
    Windows("Tabla_Conversion_MH_Lims_2.xlsx").Activate
    Sheets("Sheet1").Select
    
    ' Selecciona rango de datos desde A2:G2 hasta el final
    Range("A2:G2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Set importrange = Selection
    
    ' Aplica filtro por m�todo (columna H)
    Columns("A:H").Select
    Selection.AutoFilter
    Range("H1").Select
    ActiveSheet.Range("$A$1:$H$30000").AutoFilter Field:=8, Criteria1:=metodo
    
    ' Copia solo las celdas visibles despu�s del filtrado
    importrange.SpecialCells(xlCellTypeVisible).Select
    Selection.Copy
    
    ' --- PEGADO DE DATOS ---
    ' Vuelve al libro original y pega los valores
    Windows(nombre_excel).Activate
    Sheets("InformeFinal").Range("Q3").PasteSpecial Paste:=xlPasteValues
    
    ' Ajusta el tama�o de la tabla existente para incluir los nuevos datos
    Range("Q2:W2").Select
    Range(Selection, Selection.End(xlDown)).Select
    ActiveSheet.ListObjects("matrixmhlims").Resize Selection
    
    ' --- LIMPIEZA FINAL ---
    ' Cierra el archivo de par�metros sin guardar cambios
    Application.CutCopyMode = False
    Windows("Tabla_Conversion_MH_Lims_2.xlsx").Activate
    Range("A1").Select
    Selection.AutoFilter
    ActiveWindow.Close False
    
    ' Vuelve a proteger la hoja y reactiva actualizaciones
    Sheets("InformeFinal").Protect Password:="0000"
    Sheets("Samples").Activate
    Application.ScreenUpdating = True
End Sub




