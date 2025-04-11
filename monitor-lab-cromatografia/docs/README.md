# Monitor de Errores para Laboratorio de Cromatografía

Aplicación para monitorear equipos de laboratorio (GC-MS, LC-MS) y detectar errores en tiempo real.

## Tecnologías Usadas
- Python 3
- Tkinter (GUI)
- Pandas (procesamiento de datos)
- BeautifulSoup (análisis de logs HTML)

## Funcionalidades
- Monitoreo automático de secuencias.
- Detección de errores en equipos.
- Estimación de tiempos de finalización.
- Exportación de resultados a Excel.

## Instalación
1. Clona el repositorio:
   ```bash
   git clone https://github.com/tu-usuario/monitor-lab-cromatografia.git

Descripción General

La aplicación lista los equipos de inyección del laboratorio dividido por secciones. Al lanzarla, 
se pueden observar los nombres de los equipos junto a su estado (Inyectando, Secuencia Finalizada o con Error) 
y el tiempo restante hasta que el equipo acabe de inyectar las secuencias programadas.

Cada 15 minutos de manera automática o cuando se pulse el botón "Actualizar ahora", la información en la ventana de 
visualización que se dejará abierta en segundo plano será actualizada. Si en alguna actualización un equipo acaba la 
secuencia de inyección o se para por algún error se abrirá una ventana emergente con prioridad alta, para asegurar 
que se visualice, que informará de este cambio al analista.

El objetivo de esta aplicación es reducir tiempos muertos por despiste o desconocimiento de los errores, 
ya que los analistas se ecuentran en una sala diferente a los equipos y no es óptimo estar revisando el estado de cada 
uno de ellos cada hora y por lo tanto aumentar la producción.