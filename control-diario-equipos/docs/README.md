# Monitor de Equipos de Cromatografía GC-MS/LC-MS

Aplicación de escritorio para monitorear el estado de equipos de laboratorio (GC-MS, LC-MS) en tiempo real, detectando inyecciones, errores y muestras saltadas.

## 🛠 Tecnologías
- **Python 3.x** (Lógica principal)
- **Tkinter** (Interfaz gráfica)
- **Pandas** (Procesamiento de datos)
- **Regex** (Análisis de patrones de archivos)
- **XML/BeautifulSoup** (Procesamiento de logs)

## 🌟 Funcionalidades clave
- ✅ Monitoreo automatizado de secuencias en equipos:
  - Semivolátiles
  - Volátiles
  - Fenoles
  - Hidrocarburos
  - Twister (SPME)
- 📊 Detección de:
  - Inyecciones completadas
  - Muestras saltadas (con alertas)
  - Errores en logs (`mslogbk.htm`)
- ⏱ Estimación de tiempos de finalización
- 📤 Exportación de resultados a Excel (compatible con macros)

## 📦 Instalación
1. Clona el repositorio:
   ```bash
   git clone https://github.com/tu-usuario/monitor-cromatografia.git
   cd monitor-cromatografia

Descripción General
La rutina consistía en cada día al llegar al laboratorio verificar el estado de los equipos, 
contar las inyecciones a mano que había hecho cada uno de tus equipos (en el laboratorio hay 30 equipos en total) 
viendo la hora de la inyección y teniendo en cuenta que en un día las inyecciones pueden estar divididas en 3 
secuencias diferentes. La cantidad de inyecciones se necesita para el departamento de mejora continua.

Tras ver que esta tarea consumía unos 15 minutos diarios aproxidamante por sección, desarrollé una aplicación que, 
tras indicar tu sección y la fecha que se desea revisar, te devuelve en cuestión de segundos una lista con los equipos correspondientes, 
la cantidad de inyecciones y el estado de estos.

Pese a ser algo que parece simple, a largo plazo supone un ahorro de tiempo sustancial por parte del personal cualificado 
que se puede dedicar a otras tareas relacionas con la producción.