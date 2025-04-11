import os
import datetime
import tkinter as tk
from tkinter import ttk, messagebox
import calendar
import locale
import re
import xml.etree.ElementTree as ET
import traceback

# Configuración regional para fechas en español
locale.setlocale(locale.LC_TIME, 'es_ES')

def obtener_carpeta_por_dia(fecha, ruta_mes):
    """
    Busca una carpeta de secuencia por día en una ruta de mes.
    
    Args:
        fecha (datetime): Fecha a buscar (ej: datetime(2024, 6, 25)).
        ruta_mes (str): Ruta base del mes (ej: "Ruta\06.Junio").
    
    Returns:
        str: Ruta completa de la carpeta del día si existe, None si no.
    """
    dia_formato = fecha.strftime("Secuencia %Y-%m-%d")
    for carpeta in os.listdir(ruta_mes):
        if carpeta.startswith(dia_formato):
            return os.path.join(ruta_mes, carpeta)
    return None
def contar_inyecciones_y_revisar_equipo(ruta_base, fecha, equipo):
    """
    Cuenta inyecciones y verifica el estado de un equipo en un día específico.
    
    Args:
        ruta_base (str): Ruta base donde buscar (ej: "Ruta\2024").
        fecha (str): Fecha en formato "DD-MM-YYYY".
        equipo (str): Identificador del equipo (ej: "CS-4078").
    
    Returns:
        tuple: (conteo_inyecciones, ultima_hora, equipo_parado, muestras_saltadas, error)
            - conteo_inyecciones: Número de inyecciones encontradas.
            - ultima_hora: Hora de la última inyección (datetime).
            - equipo_parado: True si el equipo está inactivo.
            - muestras_saltadas: Lista de muestras omitidas (solo Volátiles).
            - error: Mensaje de error si falla.
    """
    try:
        # 1. Preparación de rutas y rangos de tiempo
        dia_seleccionado = datetime.datetime.strptime(fecha, "%d-%m-%Y")
        mes_nombre = f"{dia_seleccionado.month:02d}.{calendar.month_name[dia_seleccionado.month].capitalize()}"
        ruta_mes = os.path.join(ruta_base, mes_nombre)

        # 2. Búsqueda de carpeta del día (con retroceso si no se encuentra)
        if os.path.exists(ruta_mes):
            carpeta_dia = obtener_carpeta_por_dia(dia_seleccionado, ruta_mes)
        else:
            carpeta_dia = None
        dias_busqueda = 1
        while not carpeta_dia:
            # Buscar el día inmediatamente anterior
            dia_anterior1 = dia_seleccionado - datetime.timedelta(days=dias_busqueda)
            mes_anterior_nombre = f"{dia_anterior1.month:02d}.{calendar.month_name[dia_anterior1.month].capitalize()}"

            # Verificar si hay un cambio de año
            if dia_anterior1.year != dia_seleccionado.year:
                ruta_base_anterior = ruta_base.replace(fecha[-4:], f"{int(fecha[-4:])-1}")
            else:
                ruta_base_anterior = ruta_base

            ruta_mes_anterior = os.path.join(ruta_base_anterior, mes_anterior_nombre)

            if not os.path.exists(ruta_mes_anterior):
                raise FileNotFoundError(f"No se encontró la carpeta del mes anterior: {mes_anterior_nombre}")

            carpeta_dia = obtener_carpeta_por_dia(dia_anterior1, ruta_mes_anterior)
            if not carpeta_dia:
                dias_busqueda += 1  # Incrementar los días a retroceder
                if dias_busqueda > 10:  # Limitar la búsqueda a los primeros 10 días anteriores
                    raise FileNotFoundError("No se encontraron carpetas para el día seleccionado ni para los días anteriores.")
                
        if not os.path.exists(ruta_mes) and not carpeta_dia:
            raise FileNotFoundError(f"No se encontró la carpeta del mes: {mes_nombre}")

        # 3. Definir rango horario (7 AM a 7 AM del día siguiente)
        inicio_rango = datetime.datetime.combine(dia_seleccionado, datetime.time(7, 0))
        fin_rango = inicio_rango + datetime.timedelta(days=1)

        # 4. Contar inyecciones y detectar muestras saltadas
        conteo_inyecciones = 0
        ultima_inyeccion_hora = None
        equipo_parado = False
        patron_fecha = re.compile(r'^\d{2}-\d{2}-\d{2}')
        ultima_carpeta_inyeccion = None

        # Procesar carpetas del día, día anterior y siguiente, en el caso de haberlos
        carpetas_a_revisar = [carpeta_dia]

        dias_busqueda2 = 1
        dia_anterior2 = dia_seleccionado - datetime.timedelta(days=dias_busqueda2)
        mes_anterior_nombre2 = f"{dia_anterior2.month:02d}.{calendar.month_name[dia_anterior2.month].capitalize()}"

        # Verificar si hay un cambio de año
        if dia_anterior2.year != dia_seleccionado.year:
            ruta_base_anterior2 = ruta_base.replace(fecha[-4:], f"{int(fecha[-4:])-1}")
        else:
            ruta_base_anterior2 = ruta_base

        ruta_mes_anterior2 = os.path.join(ruta_base_anterior2, mes_anterior_nombre2)
        carpeta_dia_anterior = obtener_carpeta_por_dia(dia_anterior2, ruta_mes_anterior2)
        while not carpeta_dia_anterior and dias_busqueda2 < 5:
            try:
                dia_anterior2 = dia_seleccionado - datetime.timedelta(days=dias_busqueda2)
                mes_anterior_nombre2 = f"{dia_anterior2.month:02d}.{calendar.month_name[dia_anterior2.month].capitalize()}"

                # Verificar si hay un cambio de año
                if dia_anterior2.year != dia_seleccionado.year:
                    ruta_base_anterior2 = ruta_base.replace(fecha[-4:], f"{int(fecha[-4:])-1}")
                else:
                    ruta_base_anterior2 = ruta_base

                ruta_mes_anterior2 = os.path.join(ruta_base_anterior2, mes_anterior_nombre2)
                carpeta_dia_anterior = obtener_carpeta_por_dia(dia_anterior2, ruta_mes_anterior2)
                if not carpeta_dia_anterior:
                    dias_busqueda2 += 1  # Incrementar los días a retroceder
            except:
                carpeta_dia_anterior = carpeta_dia

        if carpeta_dia_anterior != carpeta_dia:
            carpetas_a_revisar.insert(0, carpeta_dia_anterior)

        dia_siguiente = dia_seleccionado + datetime.timedelta(days=1)
        mes_siguiente_nombre = f"{dia_siguiente.month:02d}.{calendar.month_name[dia_siguiente.month].capitalize()}"

        # Verificar si hay un cambio de año
        if dia_siguiente.year != dia_seleccionado.year:
            ruta_base_siguiente = ruta_base.replace(fecha[-4:], f"{int(fecha[-4:])+1}")
        else:
            ruta_base_siguiente = ruta_base

        ruta_mes_siguiente = os.path.join(ruta_base_siguiente, mes_siguiente_nombre)
        if os.path.exists(ruta_mes_siguiente):
            carpeta_dia_siguiente = obtener_carpeta_por_dia(dia_siguiente, ruta_mes_siguiente)
            if carpeta_dia_siguiente:
                carpetas_a_revisar.append(carpeta_dia_siguiente)

        muestras_saltadas = []
        for carpeta in carpetas_a_revisar:
            for subcarpeta in os.listdir(carpeta):
                if patron_fecha.match(subcarpeta):
                    carpeta_inyeccion = os.path.join(carpeta, subcarpeta)
                    # Obtener la hora de creación de la carpeta
                    hora_creacion = datetime.datetime.fromtimestamp(os.path.getctime(carpeta_inyeccion))
                    if inicio_rango <= hora_creacion < fin_rango:
                        conteo_inyecciones += 1
                        if ultima_inyeccion_hora is None or hora_creacion > ultima_inyeccion_hora:
                            ultima_inyeccion_hora = hora_creacion
                            ultima_carpeta_inyeccion = carpeta_inyeccion
                        # Comprobar si se ha saltado inyecciones (solo para volátiles)
                        archivos = os.listdir(carpeta_inyeccion)
                        if len(archivos) < 7:
                            if hora_creacion < datetime.datetime.now() - datetime.timedelta(minutes=27):
                                conteo_inyecciones -= 1
                                match_sample = re.search(rf'_(SAMPLE \d+)', subcarpeta, re.IGNORECASE)
                                if match_sample:
                                    muestras_saltadas.append(match_sample.group(1))

        # 5. Verificar si el equipo está parado (menos de 8 archivos en última carpeta)
        if ultima_carpeta_inyeccion:
            archivos = os.listdir(ultima_carpeta_inyeccion)
            if len(archivos) < 8:
                equipo_parado = True

        return conteo_inyecciones, ultima_inyeccion_hora, equipo_parado, muestras_saltadas, None  # Sin error

    except Exception as e:
        return None, None, None, None, str(e)  # Capturamos el error y lo retornamos
        
def contar_hidrocarburos(ruta_base, fecha):
    """
    Versión especializada para hidrocarburos (patrones de archivos distintos).
    
    Args:
        ruta_base (str): Ruta base del año (ej: "Ruta\2024").
        fecha (str): Fecha en formato "DD-MM-YYYY".
    
    Returns:
        tuple: (conteo, ultima_hora, error)
    """
    try:
        # Formateo de fecha para localizar las carpetas
        dia_seleccionado = datetime.datetime.strptime(fecha, "%d-%m-%Y")
        mes_nombre = f"{dia_seleccionado.month:02d}.{calendar.month_name[dia_seleccionado.month].capitalize()}"
        ruta_mes = os.path.join(ruta_base, mes_nombre)

        # Función auxiliar para encontrar carpetas por fecha
        def obtener_carpeta_por_dia(fecha, ruta_mes):
            dia_formato = fecha.strftime("Secuencia %Y-%m-%d")
            for carpeta in os.listdir(ruta_mes):
                if carpeta.startswith(dia_formato):
                    return os.path.join(ruta_mes, carpeta)
            return None

        # Buscar carpeta del día seleccionado
        if os.path.exists(ruta_mes):
            carpeta_dia = obtener_carpeta_por_dia(dia_seleccionado, ruta_mes)
        else:
            carpeta_dia = None

        dias_busqueda = 1
        while not carpeta_dia:
            # Buscar el día inmediatamente anterior
            dia_anterior = dia_seleccionado - datetime.timedelta(days=dias_busqueda)
            mes_anterior_nombre = f"{dia_anterior.month:02d}.{calendar.month_name[dia_anterior.month].capitalize()}"

            # Verificar si hay un cambio de año
            if dia_anterior.year != dia_seleccionado.year:
                ruta_base_anterior = ruta_base.replace(fecha[-4:], f"{int(fecha[-4:])-1}")
            else:
                ruta_base_anterior = ruta_base

            ruta_mes_anterior = os.path.join(ruta_base_anterior, mes_anterior_nombre)

            if not os.path.exists(ruta_mes_anterior):
                raise FileNotFoundError(f"No se encontró la carpeta del mes anterior: {mes_anterior_nombre}")

            carpeta_dia = obtener_carpeta_por_dia(dia_anterior, ruta_mes_anterior)
            if not carpeta_dia:
                dias_busqueda += 1  # Incrementar los días a retroceder
                if dias_busqueda > 10:  # Limitar la búsqueda a los primeros 10 días anteriores
                    raise FileNotFoundError("No se encontraron carpetas para el día seleccionado ni para los días anteriores.")

        if not os.path.exists(ruta_mes) and not carpeta_dia:
            raise FileNotFoundError(f"No se encontró la carpeta del mes: {mes_nombre}")

        # Determinar el rango de búsqueda (7 AM a 7 AM del día siguiente)
        inicio_rango = datetime.datetime.combine(dia_seleccionado, datetime.time(7, 0))
        fin_rango = inicio_rango + datetime.timedelta(days=1)

        conteo = 0
        ultima_inyeccion_hora = None

        # Buscar carpetas que sigan el formato DD-MM-YY_TPH
        patron_tph = re.compile(r'^\d{2}-\d{2}-\d{2}_TPH.*$')
        patron_set_cg = re.compile(r'^(Set_CG-\d{3}-a_(Front|Back)_\d{2}-\d{2}-\d{2}).*$')
        patron_desglose = re.compile(r'^Set_Desglose_(Front|Back)_\d{2}-\d{2}-\d{2}.*$')
        patron_dx = re.compile(r'^\d{2}-\d{2}-\d{2}_TPH.*\.dx$')

        carpetas_a_revisar = [carpeta_dia]

        # Buscar carpeta del día anterior (solo si existe)
        dia_anterior = dia_seleccionado - datetime.timedelta(days=1)
        mes_anterior_nombre = f"{dia_anterior.month:02d}.{calendar.month_name[dia_anterior.month].capitalize()}"

        # Verificar si hay un cambio de año
        if dia_anterior.year != dia_seleccionado.year:
            ruta_base_anterior = ruta_base.replace(fecha[-4:], f"{int(fecha[-4:])-1}")
        else:
            ruta_base_anterior = ruta_base

        ruta_mes_anterior = os.path.join(ruta_base_anterior, mes_anterior_nombre)
        if os.path.exists(ruta_mes_anterior):
            carpeta_dia_anterior = obtener_carpeta_por_dia(dia_anterior, ruta_mes_anterior)
            if carpeta_dia_anterior and carpeta_dia_anterior != carpeta_dia:
                carpetas_a_revisar.insert(0, carpeta_dia_anterior)

        # Buscar carpeta del día siguiente (solo si existe)
        dia_siguiente = dia_seleccionado + datetime.timedelta(days=1)
        mes_siguiente_nombre = f"{dia_siguiente.month:02d}.{calendar.month_name[dia_siguiente.month].capitalize()}"

        # Verificar si hay un cambio de año
        if dia_siguiente.year != dia_seleccionado.year:
            ruta_base_siguiente = ruta_base.replace(fecha[-4:], f"{int(fecha[-4:])+1}")
        else:
            ruta_base_siguiente = ruta_base

        ruta_mes_siguiente = os.path.join(ruta_base_siguiente, mes_siguiente_nombre)
        if os.path.exists(ruta_mes_siguiente):
            carpeta_dia_siguiente = obtener_carpeta_por_dia(dia_siguiente, ruta_mes_siguiente)
            if carpeta_dia_siguiente:
                carpetas_a_revisar.append(carpeta_dia_siguiente)

        # Revisar carpetas en la carpeta del día
        for carpeta in carpetas_a_revisar:
            for subcarpeta in os.listdir(carpeta):
                ruta_carpeta_inyeccion = os.path.join(carpeta, subcarpeta)

                if patron_tph.match(subcarpeta):
                    # Verificar la hora de creación de la carpeta
                    hora_creacion = datetime.datetime.fromtimestamp(os.path.getctime(ruta_carpeta_inyeccion))
                    if inicio_rango <= hora_creacion < fin_rango:
                        conteo += 1
                        if ultima_inyeccion_hora is None or hora_creacion > ultima_inyeccion_hora:
                            ultima_inyeccion_hora = hora_creacion

                elif patron_set_cg.match(subcarpeta) or patron_desglose.match(subcarpeta):
                    # Revisar archivos dentro de estas carpetas
                    for archivo in os.listdir(ruta_carpeta_inyeccion):
                        if patron_dx.match(archivo):
                            ruta_archivo = os.path.join(ruta_carpeta_inyeccion, archivo)
                            hora_creacion = datetime.datetime.fromtimestamp(os.path.getctime(ruta_archivo))
                            if inicio_rango <= hora_creacion < fin_rango:
                                conteo += 1
                                if ultima_inyeccion_hora is None or hora_creacion > ultima_inyeccion_hora:
                                    ultima_inyeccion_hora = hora_creacion

        return conteo, ultima_inyeccion_hora, None  # No se está retornando error en este caso

    except Exception as e:
        traceback_info = traceback.format_exc()
        return None, None, traceback_info
    
# Interfaz gráfica

def main():
    """
    Configura la interfaz gráfica con:
    - Selector de fecha y sección.
    - Botón para iniciar revisión.
    - Cuadro de diálogo con resultados.
    """
    # Crear ventana principal
    ventana = tk.Tk()
    ventana.title("Revisión CG-MS")

    # Crear Notebook para las pestañas
    notebook = ttk.Notebook(ventana)
    notebook.pack(pady=10, expand=True)

    # Crear marco para la pestaña "Control Diario"
    frame_control_diario = ttk.Frame(notebook, width=400, height=280)
    frame_control_diario.pack(fill='both', expand=True)

    # Añadir pestañas al Notebook
    notebook.add(frame_control_diario, text="Control Diario")

    # Contenido de la pestaña "Control Diario"
    etiqueta_fecha = tk.Label(frame_control_diario, text="Fecha (DD-MM-YYYY):")
    etiqueta_fecha.pack(pady=5)
    entrada_fecha = tk.Entry(frame_control_diario)
    entrada_fecha.pack(pady=5)

    etiqueta_seccion = tk.Label(frame_control_diario, text="Sección:")
    etiqueta_seccion.pack(pady=5)

    opciones_seccion = ["Fenoles", "Semivol", "Volátiles", "Hidrocarburos", "Twister"]
    seccion_var = tk.StringVar()
    seccion_dropdown = tk.OptionMenu(frame_control_diario, seccion_var, *opciones_seccion)
    seccion_dropdown.pack(pady=5)

# Botón para procesar la información
    def procesar_informacion():
        """
        Ejecuta la lógica al hacer clic en el botón:
        - Valida inputs.
        - Llama a contar_inyecciones_y_revisar_equipo o contar_hidrocarburos.
        - Muestra resultados en un messagebox.
        """
        fecha = entrada_fecha.get()
        seccion = seccion_var.get()

        if not fecha or not seccion:
            messagebox.showwarning("Advertencia", "Debe ingresar todos los datos.")
            return

        if seccion == "Hidrocarburos":
            rutas = {
                "CS-4307": rf"\\d073-b73wq2rlo4\d\CDSProjects\LTM\Results\{fecha[-4:]}", 
                "CS-4719": rf"\\Desktop-5lvhtdp\d\CDSProjects\LTM2\Results\{fecha[-4:]}", 
            }
            tiempo_inyeccion = datetime.timedelta(minutes=7)
            resultados = []
            for nombre_ruta, ruta in rutas.items():
                conteo_total = 0
                ultima_inyeccion_hora = None
                mensaje_error = None
                conteo, ultima_hora, error = contar_hidrocarburos(ruta, fecha)
                mensaje = f"{nombre_ruta}:\n"
                if conteo is not None:
                    mensaje += f"Inyecciones realizadas: {conteo}\n"
                    if ultima_hora:
                        mensaje += f"Hora de la última inyección: {ultima_hora.strftime('%H:%M:%S')}\n"
                        hora_actual = datetime.datetime.now()
                        hora_7am = datetime.datetime.combine(datetime.datetime.strptime(fecha, "%d-%m-%Y") + datetime.timedelta(days=1), datetime.time(7, 0))
                        if hora_actual < (ultima_hora + tiempo_inyeccion):
                            mensaje += "El equipo está inyectando.\n"                                                 
                        elif hora_7am - ultima_hora > tiempo_inyeccion:
                            mensaje += "El equipo ha acabado de inyectar.\n"
                        else:
                            mensaje += "El equipo está funcionando correctamente.\n"
                    else:
                        mensaje += "No se encontraron inyecciones en el rango de tiempo.\n"
                if error:
                    mensaje += f"Error: {error}\n"
                resultados.append(mensaje)

        else:
            if seccion == "Semivol":
                rutas = {
                    "CS-4332": rf"\\ruta-red\d\Data file\{fecha[-4:]}", 
                    "CS-4716": rf"\\ruta-red\d\DataFile\{fecha[-4:]}", 
                    "CS-4939": rf"\\ruta-red\d (qqq-3)\Data file\{fecha[-4:]}", 
                    "CS-5144": rf"\\ruta-red\d\Data\{fecha[-4:]}"
                }
                tiempo_inyeccion = datetime.timedelta(minutes=30)
            if seccion == "Twister":
                rutas = {
                    "CS-1037": rf"\\ruta-red\d\Data File\Secuencias CGM-019-a\Secuencias {fecha[-4:]}", 
                    "CS-1342": rf"\\ruta-red\D\Secuencias CGM-019-a\Secuencias {fecha[-4:]}", 
                    "CS-1804": rf"\\ruta-red\d\Data file\Secuencias CGM-031-a\Secuencias {fecha[-4:]}", 
                    "CS-3013": rf"\\ruta-red\d (cs-3013)\Data file\Secuencias CGM-031-a\Secuencias {fecha[-4:]}",
                    "CS-4658": rf"\\ruta-red\d\Secuencias {fecha[-4:]}",
                    "CS-5002": rf"\\ruta-red\d\DATA FILE\Secuencias {fecha[-4:]}"
                }
                tiempo_inyeccion = datetime.timedelta(minutes=72)   
            if seccion == "Volátiles":
                rutas = {
                    "CS-3194": [rf"\\ruta-red\D\Agilent 3194-3195\Data\{fecha[-4:]}\HS", rf"\\ruta-red\D\Agilent 3194-3195\Data\{fecha[-4:]}\SPME"],
                    "CS-4102": rf"\\ruta-red\cs-4101-4102\Data\{fecha[-4:]}",
                    "CS-4289": [rf"\\ruta-red\d\CS-4289-MS-FID\Data\HS-FID\{fecha[-4:]}", rf"\\ruta-red\d\CS-4289-MS-FID\Data\HS-MS\{fecha[-4:]}"],
                    "CS-4714": [rf"\\ruta-red\d\CS-4714-MS-FID\Data\FID\{fecha[-4:]}", rf"\\ruta-red\d\CS-4714-MS-FID\Data\MS\{fecha[-4:]}"], 
                    "CS-4870": rf"\\ruta-red\d\CS-4870\DATA\{fecha[-4:]}",  
                    "CS-4940": rf"\\ruta-red\d\CS-4940\Data\{fecha[-4:]}",           
                    "CS-5044": rf"\\ruta-red\d\CS-5044\Data\{fecha[-4:]}",
                    "CS-5045": [rf"\\ruta-red\d\CS-5045-SPME-HS\Data\{fecha[-4:]}\HS", rf"\\ruta-red\d\CS-5045-SPME-HS\Data\{fecha[-4:]}\SPME"],
                    "CS-5142": rf"\\ruta-red\d\CS-5142-HS-MS\Data\{fecha[-4:]}"
                }
                tiempo_inyeccion = datetime.timedelta(minutes=27)
            if seccion == "Fenoles":
                rutas = {
                    "CS-4078": [rf"\\ruta-red\d\CS-4078-3195\Data\{fecha[-4:]}\CGM-036-a",rf"\\ruta-red\d\CS-4078-3195\Data\{fecha[-4:]}\CGM-038-a",rf"\\ruta-red\d\CS-4078-3195\Data\{fecha[-4:]}\CGM-020-a"], 
                    "CS-4252": rf"\\ruta-red\d\CS-4252-CG-MS\DATA\FENOLES\{fecha[-4:]}"
                }
                tiempo_inyeccion = datetime.timedelta(minutes=33)
            resultados = []
            for nombre_ruta, ruta in rutas.items():
                conteo_total = 0
                ultima_inyeccion_hora = None
                equipo_parado = False
                muestras_saltadas = []
                mensaje_error = None
                ultima_carpeta_inyeccion = None
                mensajes_estado = []

                if isinstance(ruta, list):  # Si hay varias rutas, las procesamos todas
                    for subruta in ruta:
                        conteo, ultima_hora, equipo_parado, muestras_saltadas_ruta, error = contar_inyecciones_y_revisar_equipo(subruta, fecha, seccion)
                        if conteo is not None:
                            conteo_total += conteo
                            if ultima_inyeccion_hora is None or (ultima_hora and ultima_hora > ultima_inyeccion_hora):
                                ultima_inyeccion_hora = ultima_hora
                            if equipo_parado:
                                equipo_parado = True
                            if muestras_saltadas_ruta:
                                muestras_saltadas.extend(muestras_saltadas_ruta)
                            if error:
                                mensaje_error = error

                    # Crear un mensaje consolidado para todas las subrutas
                    mensaje = f"{nombre_ruta}:\n"
                    mensaje += f"Inyecciones realizadas: {conteo_total}\n"
                    if ultima_inyeccion_hora:
                        mensaje += f"Hora de la última inyección: {ultima_inyeccion_hora.strftime('%H:%M:%S')}\n"
                        hora_actual = datetime.datetime.now()
                        hora_7am = datetime.datetime.combine(datetime.datetime.strptime(fecha, "%d-%m-%Y") + datetime.timedelta(days=1), datetime.time(7, 0))
                        if hora_actual < (ultima_inyeccion_hora + tiempo_inyeccion) and equipo_parado:
                            mensaje += "El equipo está inyectando.\n"                    
                        elif equipo_parado:
                            mensaje += "AVISO: EL EQUIPO SE HA PARADO.\n"                       
                        elif hora_actual < ultima_inyeccion_hora + tiempo_inyeccion:
                            mensaje += "El equipo está inyectando.\n"
                        elif hora_7am - ultima_inyeccion_hora > tiempo_inyeccion:
                            mensaje += "El equipo ha acabado de inyectar.\n"
                        else:
                            mensaje += "El equipo está funcionando correctamente.\n"
                        if seccion == "Volátiles" and muestras_saltadas:
                            mensaje += f"Muestras saltadas: {', '.join(muestras_saltadas)}\n"
                    else:
                        mensaje += "No se encontraron inyecciones en el rango de tiempo.\n"
                    if mensaje_error:
                        mensaje += f"Error: {mensaje_error}\n"
                    resultados.append(mensaje)

                else:
                    conteo, ultima_hora, equipo_parado, muestras_saltadas, error = contar_inyecciones_y_revisar_equipo(ruta, fecha, seccion)
                    mensaje = f"{nombre_ruta}:\n"
                    if conteo is not None:
                        mensaje += f"Inyecciones realizadas: {conteo}\n"
                        if ultima_hora:
                            mensaje += f"Hora de la última inyección: {ultima_hora.strftime('%H:%M:%S')}\n"
                            hora_actual = datetime.datetime.now()
                            hora_7am = datetime.datetime.combine(datetime.datetime.strptime(fecha, "%d-%m-%Y") + datetime.timedelta(days=1), datetime.time(7, 0))
                            if hora_actual < (ultima_hora + tiempo_inyeccion) and equipo_parado:
                                mensaje += "El equipo está inyectando.\n"                       
                            elif equipo_parado:
                                mensaje += "AVISO: EL EQUIPO SE HA PARADO.\n"                            
                            elif hora_7am - ultima_hora > tiempo_inyeccion:
                                mensaje += "El equipo ha acabado de inyectar.\n"
                            else:
                                mensaje += "El equipo está funcionando correctamente.\n"
                            if seccion == "Volátiles" and muestras_saltadas:
                                mensaje += f"Muestras saltadas: {', '.join(muestras_saltadas)}\n"
                        else:
                            mensaje += "No se encontraron inyecciones en el rango de tiempo.\n"
                    if error:
                        mensaje += f"Error: {error}\n"
                    resultados.append(mensaje)

        mensaje_final = "\n".join(resultados)
        messagebox.showinfo("Resultados", mensaje_final)

    boton_iniciar = tk.Button(frame_control_diario, text="Iniciar Revisión", command=procesar_informacion)
    boton_iniciar.pack(pady=20)

    ventana.mainloop()

if __name__ == "__main__":
    main()