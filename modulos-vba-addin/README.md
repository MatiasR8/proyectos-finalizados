# Excel Automation Portfolio

Este repositorio contiene una colección de macros VBA para Excel desarrolladas para automatizar procesos en un laboratorio de análisis, específicamente para la gestión de muestras, calibraciones y generación de reportes.

## 📋 Descripción General

El proyecto consiste en un sistema integrado de macros que:
- Gestiona la importación/exportación de datos desde/hacia LIMS (Sistema de Gestión de Laboratorio)
- Automatiza la generación de reportes en PDF
- Controla procesos de calibración y validación
- Maneja datos de muestras, blancos y controles de calidad

## 🛠 Macros Principales

### 1. Gestión de Muestras
- `Guardar_datos`: Exporta muestras al LIMS con validaciones
- `BuscarYActualizarTodo`: Actualiza datos gemelos entre hojas
- `Blancos`: Importa datos de blancos según método analítico

### 2. Procesos de Calibración
- `Calibrarcalibracion`: Exporta criterios de calibración a PDF
- `Importar_calibracion`: Importa parámetros de calibración

### 3. Generación de Reportes
- `ExportReport`: Exporta reportes a PDF con control de versiones
- `ExportarQC`: Exporta controles de calidad

### 4. Integración con LIMS
- `Importar_archivo_lims`: Importa datos desde el LIMS
- `Importar_Matrix_parametros_MH_lims`: Importa matrices de parámetros

### 5. Validaciones Especiales
- `Plaguicidas`: Identifica muestras con parámetros no subidos
- `c5c40`: Maneja controles específicos C5-C40

## 🧩 Estructura del Código

Todas las macros comparten características comunes:
- Manejo estructurado de errores
- Validación de rutas y directorios
- Protección/desprotección segura de hojas
- Optimización de rendimiento (desactivación de actualizaciones de pantalla)
- Interacción con usuario mediante mensajes

## 💻 Requisitos Técnicos

- Microsoft Excel 2016 o superior
- Habilitar macros al abrir el archivo
- Permisos de escritura en las rutas especificadas
- Estructura de carpetas según configuración en hoja "Samples"

## 🚀 Cómo Usar

1. Abrir el archivo Excel habilitando macros
2. Configurar rutas base en hoja "Samples":
   - `rutalims`: Ruta de archivos LIMS
   - `rutaexportreport`: Ruta para reportes PDF
   - `rutaparametros`: Ruta de parámetros
3. Ejecutar macros desde:
   - Botones asignados en las hojas
   - Menú Developer → Macros

## 📌 Notas Importantes

- Todas las contraseñas de protección son "0000" (fines demostrativos)
- Los nombres de archivo generados eliminan automáticamente caracteres especiales
- Las macros están diseñadas para un flujo de trabajo específico de laboratorio


---

✨ Desarrollado como parte de mi portfolio técnico - Matías Rodríguez - 2025