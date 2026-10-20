"""
============================================================================
 Lector de "la manita" (por Bluetooth directo a PEPE) -> Excel
============================================================================

Que hace:
  - Abre un puerto serie: directamente el puerto COM que Windows le asigna a PEPE
    cuando lo emparejas por Bluetooth -- el script no necesita saber cual
    de los dos es, ambos le llegan como un puerto serie normal.
  - Este script interpreta cada linea que llega, buscando pares tipo
    "ETIQUETA:VALOR" (por ejemplo "AX:120 AY:45 AZ:-30 GX:.. ANG:.. F1:..",
    sin importar el separador exacto entre pares: espacio, coma, tab, ";"),
    y va llenando un archivo de Excel con una columna por cada etiqueta
    (ax, ay, az, gx, gy, gz, GX, GY, D0-D4) y una fila por cada lectura.
  - Las columnas van en un orden FIJO (ver COLUMNAS_ORDEN mas abajo en el
    script): D0, D1, D2, D3, D4, ax, ay, az, gx, gy, gz, GX, GY -- sin
    importar en que orden lleguen los datos de la manita.

MEDICIONES POR SEPARADO (un archivo nuevo por cada toma):
  Cada vez que enciendes la manita, empieza a llegar un chorro continuo de
  lecturas; cuando la apagas, ese chorro se corta. Este script usa ese
  "silencio" (ver SIN_DATOS_CIERRA_MEDICION_S mas abajo: si pasan varios
  segundos sin ninguna lectura real) para saber cuando termina una
  medicion, y la siguiente lectura que llegue despues de un silencio
  arranca la siguiente medicion:

    - En cuanto llega la PRIMERA lectura real de una medicion nueva, se
      prepara un archivo Excel NUEVO (con fecha y hora en el nombre, por
      ejemplo datos_manita_20260909_153000.xlsx) y el cronometro interno
      arranca en 0 -- ya no se guarda la fecha/hora de cada lectura, sino
      los SEGUNDOS transcurridos desde esa primera lectura (columna
      "t_segundos"). Tambien genera un archivo de respaldo de datos, pero
      este no tiene gráficas.
    - Cuando se detecta el silencio (apagaste la manita) o detienes el
      script (Ctrl+C), ese archivo se cierra y se le agregan las 4 hojas
      de graficas (ver abajo). La siguiente vez que enciendas la manita
      se abre un archivo distinto, sin mezclar mediciones.
    - Si el puerto llega a fallar de plano (por ejemplo se te apago la 
      manita a medio camino del todo y Windows cierra el puerto), el script 
      no se cierra: cierra la medicion que tenia a medias y se queda 
      reintentando abrir elpuerto solo, hasta que vuelvas a encender la manita.
      Por lo generar para hacer una nueva medicion se necesita correr el 
      programa de nuevo.

GRAFICAS:
  Cada archivo de una medicion queda con 5 hojas en total:
    1. "Datos"                -> la tabla completa (t_segundos + columnas)
    2. "Grafica_Dedos"        -> D0, D1, D2, D3, D4  (una linea de color por dedo)
    3. "Grafica_Acelerometro" -> ax, ay, az
    4. "Grafica_Giroscopio"   -> gx, gy, gz
    5. "Grafica_Angulos"      -> GX, GY
  Todas son graficas de linea 2D, con el tiempo (segundos) en el eje
  horizontal (abscisas) y el valor de cada variable en el eje vertical
  (ordenadas), agrandadas para que se vean bien.

Requisitos (instalar una sola vez):
    pip install pyserial openpyxl

Antes de correrlo:
  Directo a PEPE por Bluetooth :
    1. Empareja PEPE en el Bluetooth de Windows.
    2. Busca en "Mas opciones de Bluetooth" -> pestana "Puertos COM" el
       puerto COM saliente que Windows le asigno a PEPE.
    3. Pon ese puerto en SERIAL_PORT mas abajo (por ejemplo "COM3") --
       aqui NO sirve "auto", porque un puerto Bluetooth no se distingue
       por el mismo chip USB-serie que una ESP32.
    4. Enciende la manita ANTES de correr el script (aqui es Windows/
       Python quien inicia la conexion al abrir el puerto, no hay nada
       reintentando solo del otro lado como hace la ESP32).
    5. Calibra la manita: la correcta calibracion es haciendo primero un puño 
       (se encenderán los LEDs azules de la manita y luego se apagarán), despues
       abrir lentamente (aprox 2 segundos) la palma hasta tener los dedos
       extendidos (se volveran a encender los LEDs azules y luego se apagarán). 
       Asi la manita sabrá los valores limites de movimiento. Ahora la manita ya
       esta calibrada.
    6. Corre:  python manita_a_excel.py
    7. Enciende/apaga la manita las veces que necesites -- cada ciclo
       genera su propio archivo de Excel, ya con las graficas incluidas y con un
       archivo de respaldo.
    8. Para detener el script por completo, Ctrl+C.

NOTA sobre el formato de los datos:
  El parser esta ajustado al formato real de tu manita (visto en la
  terminal Bluetooth), por ejemplo:

    Accel: ax:-0.05 ay:0.05 az:1.01 Gyro: gx:-50.59 gy:157.20 gz:76.00
    Angles: GX:-2.76 GY:2.99 Fingers: D0:500.0 D1:500.0 D2:500.0 D3:500.0
    D4:500.0

  o el formato compacto del firmware v2 (lehand_30hz_elbueno.ino actualizado), que
  ademas trae "t" = milisegundos de la manita:

    t:15321 ax:-0.05 ay:0.05 az:1.01 gx:-51 gy:157 gz:76 GX:-2.8 GY:3.0
    D0:500 D1:500 D2:500 D3:500 D4:500

  Funciona con pares "etiqueta:valor" (o "etiqueta=valor"), sin importar
  cuantos espacios/comas/tabs haya entre ellos, y respeta las palabras
  "Accel:", "Gyro:", "Angles:" y "Fingers:" (se ignoran solas, no generan
  columnas). Las etiquetas se guardan tal cual las manda la manita,
  RESPETANDO mayusculas/minusculas -- esto es importante porque tu modulo
  usa "gx"/"gy" para el giroscopio y "GX"/"GY" para el angulo ya
  calculado; si no se respetara el caso, se pisarian entre si.

  Columnas esperadas con tu formato: ax, ay, az (acelerometro), gx, gy, gz
  (giroscopio), GX, GY (angulo calculado), D0-D4 (los 5 dedos).

============================================================================
"""

import csv
import re
import sys
import time
from datetime import datetime
from pathlib import Path

import serial
from openpyxl import Workbook
from openpyxl.chart import LineChart, Reference

# ============================== CONFIGURACION ===============================

# Puerto serie a abrir:
# Conectando directo a PEPE por Bluetooth,
# pon aqui el puerto COM que le asigno Windows (ver "Puertos COM" en
# "Mas opciones de Bluetooth") 
#   Windows:      "COM5", "COM3", etc.
#   macOS:        "/dev/cu.usbserial-0001"  o  "/dev/cu.SLAB_USBtoUART"
#   Linux:        "/dev/ttyUSB0"
SERIAL_PORT = "COM3"

BAUD_RATE = 115200  # debe coincidir con Serial.begin(...) en el .ino; ahorita ya 
#estan alineados la manita y este archivo al mismo baud rate

CARPETA_SALIDA = Path(__file__).resolve().parent
ARCHIVO_LOG_CRUDO = CARPETA_SALIDA / "manita_log_crudo.txt"
# Ya NO hay un unico ARCHIVO_EXCEL/ARCHIVO_CSV fijo: cada medicion (cada vez
# que se enciende la manita) genera su propio nombre de archivo con fecha y
# hora, ver "nuevo_nombre_archivo()" mas abajo.

GUARDAR_EXCEL_CADA_S = 3.0       # cada cuantos segundos se re-guarda el .xlsx
GUARDAR_EXCEL_CADA_N_FILAS = 300  # o cada cuantas filas nuevas, lo que ocurra primero
# A ~30 lecturas/seg, cada guardado reescribe TODO el archivo desde cero, y
# el archivo va creciendo -- por eso el intervalo se subio de 2s a 3s (a 20
# filas/1000ms se guardaba practicamente en cada vuelta del loop, lo cual es
# innecesario y con sesiones largas lo iria haciendo mas lento). Estos
# guardados periodicos son solo de respaldo (hoja "Datos", sin graficas);
# las graficas se agregan una sola vez, al cerrar la medicion (ver
# finalizar_medicion), para no repetir ese trabajo en cada guardado.

# =============================================================================

# Encuentra pares "ETIQUETA:VALOR" o "ETIQUETA=VALOR", sin importar que los
# separe entre si (espacio, coma, tab, punto y coma, etc).
PATRON_PAR = re.compile(r"([A-Za-z_][A-Za-z0-9_]*)\s*[:=]\s*(-?\d+(?:\.\d+)?)")

# Unicas etiquetas que realmente manda la manita (confirmado viendo la
# terminal Bluetooth). Cualquier otra cosa que llegue por el puerto serie
# se ignora para la hoja de calculo -- en particular, el "banner" de
# arranque que la propia ESP32 imprime sola cada vez que se resetea
# (rst:0x1, boot:0x13, configsip:0, SPIWP:0xee, clk_drv:0x00, load:...,
# len:..., etc.) tambien tiene forma de "etiqueta:valor" y sin este filtro
# se cuela como si fueran columnas de datos.
CAMPOS_MANITA = {
    "t",                # milisegundos de la manita desde que se encendio (firmware v2)
    "ax", "ay", "az",   # acelerometro
    "gx", "gy", "gz",   # giroscopio (velocidad angular)
    "GX", "GY",         # angulo ya calculado (mayusculas, distinto de gx/gy)
    "D0", "D1", "D2", "D3", "D4",  # flexion de los 5 dedos
}
# Nota sobre "t": NO se guarda como columna propia; se usa para calcular
# t_segundos con el reloj de la propia manita (el momento REAL en que se
# tomo cada lectura), en vez de la hora a la que la linea llego a la
# computadora -- el Bluetooth de Windows entrega las lineas "a rachas"
# (varias juntas en el mismo milisegundo, luego un hueco), y eso hacia
# que el tiempo guardado no fuera confiable. Si la manita tiene cargado un
# firmware viejo que no manda "t", se usa la hora de llegada como antes.

# Orden FIJO en que deben quedar las columnas en el Excel/CSV:
# D0-D4 primero, luego acelerometro, giroscopio y angulos).
# Este mismo orden es el que usan las graficas para saber en que columna de
# la hoja "Datos" esta cada variable (ver GRAFICAS_A_GENERAR mas abajo), asi
# que si cambias este orden tambien hay que revisar esa lista.
COLUMNAS_ORDEN = [
    "D0", "D1", "D2", "D3", "D4",
    "ax", "ay", "az",
    "gx", "gy", "gz",
    "GX", "GY",
]

# Que graficas generar por cada medicion: (nombre_hoja, titulo, etiquetas).
# "etiquetas" son las columnas de COLUMNAS_ORDEN que van juntas en esa
# grafica (cada una sale como una linea de color distinto).
GRAFICAS_A_GENERAR = [
    ("Grafica_Dedos", "Dedos (D0-D4) vs tiempo", ["D0", "D1", "D2", "D3", "D4"]),
    ("Grafica_Acelerometro", "Acelerometro (ax, ay, az) vs tiempo", ["ax", "ay", "az"]),
    ("Grafica_Giroscopio", "Giroscopio (gx, gy, gz) vs tiempo", ["gx", "gy", "gz"]),
    ("Grafica_Angulos", "Angulos (GX, GY) vs tiempo", ["GX", "GY"]),
]

# Si pasan mas de este numero de segundos sin recibir NINGUNA lectura real
# de la manita, se asume que se apago (o se desconecto el Bluetooth) y se
# cierra la medicion en curso -- la siguiente lectura que llegue despues
# abre una medicion (y archivo) nueva. No lo bajes demasiado: a veces el
# Bluetooth tiene hipos de un segundo sin que la manita se haya apagado.
SIN_DATOS_CIERRA_MEDICION_S = 3.0

# Si se usa la conexion DIRECTA a PEPE y el puerto serie falla
# de plano (Windows lo cierra porque PEPE se desconecto), cada cuantos
# segundos se reintenta volver a abrirlo, en vez de cerrar el script.
REINTENTO_APERTURA_PUERTO_S = 2.0


def parsear_linea(linea: str) -> dict:
    """Convierte una linea cruda tipo 'ax:-0.05 ay:0.05 GX:-2.76' en un dict.

    OJO: las etiquetas se guardan EXACTAMENTE con las mayusculas/minusculas
    con que las manda la manita, sin normalizar. Es a proposito: tu modulo
    distingue "gx"/"gy" (velocidad angular del giroscopio) de "GX"/"GY"
    (angulo ya calculado). Si aqui se pasara todo a minusculas, "gx" y "GX"
    (o "gy" y "GY") se pisarian entre si y se perderia uno de los dos datos.
    """
    pares = PATRON_PAR.findall(linea)
    datos = {}
    for etiqueta, valor in pares:
        etiqueta = etiqueta.strip()
        if etiqueta not in CAMPOS_MANITA:
            continue  # ruido (banner de arranque de la ESP32, etc.): se ignora
        datos[etiqueta] = float(valor) if "." in valor else int(valor)
    return datos


def nuevo_nombre_archivo() -> tuple:
    """Genera el nombre (Excel, CSV) de una medicion nueva, con fecha/hora.

    Si dos mediciones arrancaran dentro del mismo segundo (poco probable,
    pero posible si se enciende/apaga la manita muy rapido), se le agrega
    un sufijo numerico para no pisar el archivo de la medicion anterior.
    """
    marca = datetime.now().strftime("%Y%m%d_%H%M%S")
    base = f"datos_manita_{marca}"
    sufijo = ""
    contador = 1
    while (CARPETA_SALIDA / f"{base}{sufijo}.xlsx").exists():
        contador += 1
        sufijo = f"_{contador}"
    archivo_excel = CARPETA_SALIDA / f"{base}{sufijo}.xlsx"
    archivo_csv = CARPETA_SALIDA / f"{base}{sufijo}_respaldo.csv"
    return archivo_excel, archivo_csv


def guardar_excel_datos(columnas: list, filas: list, archivo_excel: Path) -> None:
    """Guarda solo la hoja "Datos" (sin graficas).

    Se usa como respaldo periodico MIENTRAS se recibe una medicion (por si
    el script se cierra mal a medio camino); es rapido porque no genera
    graficas. Las 4 hojas de graficas se agregan una sola vez, al terminar
    la medicion, en finalizar_medicion().
    """
    if archivo_excel is None:
        return
    try:
        wb = Workbook()
        ws = wb.active
        ws.title = "Datos"
        encabezado = ["t_segundos"] + columnas
        ws.append(encabezado)
        for fila in filas:
            ws.append([fila.get(col, "") for col in encabezado])
        wb.save(archivo_excel)
    except PermissionError:
        print(
            f"[AVISO] No pude guardar {archivo_excel.name} porque esta abierto "
            "en Excel. Cierralo para que se pueda seguir actualizando "
            "(los datos siguen respaldandose en el CSV mientras tanto)."
        )


def guardar_csv(columnas: list, filas: list, archivo_csv: Path) -> None:
    """Reescribe el .csv de respaldo completo de la medicion actual."""
    if archivo_csv is None:
        return
    try:
        with open(archivo_csv, "w", newline="", encoding="utf-8") as f:
            escritor = csv.writer(f)
            encabezado = ["t_segundos"] + columnas
            escritor.writerow(encabezado)
            for fila in filas:
                escritor.writerow([fila.get(col, "") for col in encabezado])
    except PermissionError:
        print(
            f"[AVISO] No pude guardar {archivo_csv.name} porque esta "
            "abierto en otro programa."
        )


def agregar_hoja_grafica(wb, ws_datos, nombre_hoja: str, titulo: str, etiquetas: list, max_row: int) -> None:
    """Crea una hoja nueva con una grafica de linea 2D (ampliada).

    "etiquetas" son columnas de COLUMNAS_ORDEN (por ejemplo ["ax","ay","az"]);
    cada una se agrega como una serie (linea) de color distinto, todas en la
    misma grafica, con el tiempo (t_segundos, columna A de "Datos") en el
    eje horizontal y el valor en el eje vertical.
    """
    # Columna 1 de "Datos" es t_segundos; despues vienen COLUMNAS_ORDEN en
    # orden, asi que la columna de cada etiqueta es su indice + 2 (1 para
    # pasar de base-0 a base-1, +1 mas por la columna de t_segundos).
    indices_columna = [COLUMNAS_ORDEN.index(etiqueta) + 2 for etiqueta in etiquetas]

    ws_grafica = wb.create_sheet(nombre_hoja)

    chart = LineChart()
    chart.title = titulo
    chart.style = 2
    chart.x_axis.title = "Tiempo (s)"
    chart.y_axis.title = "Valor"
    chart.width = 28   # cm -- grafica "ampliada"
    chart.height = 14  # cm

    cats_ref = Reference(ws_datos, min_col=1, min_row=2, max_row=max_row)

    # Las etiquetas de COLUMNAS_ORDEN que interesan a esta grafica no
    # siempre son columnas consecutivas en la hoja (por ejemplo GX/GY si en
    # el futuro cambiara el orden), asi que se agregan una por una en vez
    # de asumir un solo rango continuo.
    for col in indices_columna:
        datos_ref = Reference(ws_datos, min_col=col, max_col=col, min_row=1, max_row=max_row)
        chart.add_data(datos_ref, titles_from_data=True)

    chart.set_categories(cats_ref)
    ws_grafica.add_chart(chart, "A1")


def resumen_calidad(filas: list) -> str:
    """Resumen rapido para saber si una medicion salio "sana".

    Calcula cuantas lecturas por segundo llegaron en promedio y cuantos
    "huecos" hubo (dos lecturas seguidas separadas por mas de 50 ms, o sea
    ~1.5 periodos de 33 ms: señal de que se perdio al menos una lectura).
    """
    ts = [f["t_segundos"] for f in filas]
    if len(ts) < 2 or ts[-1] <= ts[0]:
        return "muy corta para evaluar"
    tasa = (len(ts) - 1) / (ts[-1] - ts[0])
    huecos = sum(1 for a, b in zip(ts, ts[1:]) if b - a > 0.05)
    return f"{tasa:.1f} lecturas/s en promedio, {huecos} lectura(s) perdida(s) (huecos > 50 ms)"


def finalizar_medicion(columnas: list, filas: list, archivo_excel: Path, archivo_csv: Path) -> None:
    """Guarda la version FINAL de una medicion: hoja "Datos" + 4 graficas.

    Se llama una sola vez, cuando la medicion termina (se detecto que se
    perdio la conexion con la manita, o se detuvo el script).
    """
    if archivo_excel is None or not filas:
        return

    try:
        wb = Workbook()
        ws_datos = wb.active
        ws_datos.title = "Datos"
        encabezado = ["t_segundos"] + columnas
        ws_datos.append(encabezado)
        for fila in filas:
            ws_datos.append([fila.get(col, "") for col in encabezado])
        max_row = 1 + len(filas)  # +1 por el encabezado

        for nombre_hoja, titulo, etiquetas in GRAFICAS_A_GENERAR:
            agregar_hoja_grafica(wb, ws_datos, nombre_hoja, titulo, etiquetas, max_row)

        wb.save(archivo_excel)
        print(f"[INFO] Medicion cerrada: {archivo_excel.name} ({len(filas)} lecturas, con graficas)")
        print(f"       Calidad: {resumen_calidad(filas)}")
    except PermissionError:
        print(
            f"[AVISO] No pude guardar {archivo_excel.name} con las graficas porque "
            "esta abierto en Excel. Cierralo y vuelve a correr el script para "
            "que se regenere con graficas (los datos ya estan a salvo en el CSV)."
        )

    guardar_csv(columnas, filas, archivo_csv)


def abrir_puerto(puerto: str, primera_vez: bool = False) -> serial.Serial:
    """Abre el puerto serie, reintentando solo si falla (no se cierra el script).

    En la primera apertura (primera_vez=True), si falla de plano se sale
    del script con un mensaje claro (probablemente el puerto esta mal
    escrito, o el dispositivo ni siquiera esta emparejado/conectado por
    USB). Despues de esa primera vez, un fallo ya no es motivo para cerrar
    el script -- solo pasa cuando se uso la conexion directa a PEPE y se
    desconecto a medio camino -- asi que se reintenta solo cada
    REINTENTO_APERTURA_PUERTO_S segundos hasta que vuelva a abrir.
    """
    intento = 0
    while True:
        try:
            return serial.Serial(puerto, BAUD_RATE, timeout=1)
        except serial.SerialException as e:
            intento += 1
            if primera_vez:
                sys.exit(f"No se pudo abrir el puerto {puerto}: {e}")
            print(
                f"[AVISO] No se pudo abrir {puerto} (intento {intento}): {e}. "
                f"Reintentando en {REINTENTO_APERTURA_PUERTO_S:.0f}s... "
                "(¿esta la manita/PEPE encendida?)"
            )
            time.sleep(REINTENTO_APERTURA_PUERTO_S)


def main() -> None:
    puerto = SERIAL_PORT 

    print(f"[INFO] Abriendo puerto {puerto} a {BAUD_RATE} baudios...")
    ser = abrir_puerto(puerto, primera_vez=True)

    time.sleep(2)  # si es una ESP32, se reinicia al abrir el puerto; le damos tiempo
    # Todo lo que se haya acumulado durante esa pausa de 2 s se descarta:
    # antes esas ~30-60 lecturas se leian todas de golpe al terminar la
    # pausa y quedaban apiladas en t = 0.00-0.03 s al inicio del Excel.
    ser.reset_input_buffer()

    columnas: list = list(COLUMNAS_ORDEN)  # orden fijo (ver COLUMNAS_ORDEN arriba)

    # Estado de la medicion actual. "filas" y los nombres de archivo se
    # reinician cada vez que se detecta una reconexion con la manita, para
    # que cada encendido/apagado quede en su propio archivo.
    filas: list = []
    t0 = None  # momento (time.time()) de la primera lectura real de la medicion actual
    archivo_excel_actual = None
    archivo_csv_actual = None
    ultima_lectura_real = None  # time.time() de la ultima lectura real recibida
    t0_manita = None            # "t" (ms, reloj de la manita) de la primera lectura de la medicion
    ultimo_t_manita = None      # "t" de la lectura anterior (para detectar si la manita se reinicio)

    ultimo_guardado = time.time()
    filas_desde_guardado = 0

    print("[INFO] Escuchando a la manita.")
    print(f"         Log crudo (para depurar el formato): {ARCHIVO_LOG_CRUDO}")
    print(
        "[INFO] Cada vez que enciendas/apagues la manita se abre/cierra un "
        "archivo de Excel distinto (con graficas), con el nombre "
        "datos_manita_<fecha>_<hora>.xlsx"
    )
    print("[INFO] Ctrl+C para detener el script por completo.\n")

    archivo_log = open(ARCHIVO_LOG_CRUDO, "a", encoding="utf-8")

    def cerrar_medicion_actual():
        """Cierra (con graficas) la medicion en curso, si tiene datos."""
        nonlocal filas, t0, archivo_excel_actual, archivo_csv_actual, ultima_lectura_real
        nonlocal t0_manita, ultimo_t_manita
        if filas:
            finalizar_medicion(columnas, filas, archivo_excel_actual, archivo_csv_actual)
        filas = []
        t0 = None
        archivo_excel_actual = None
        archivo_csv_actual = None
        ultima_lectura_real = None
        t0_manita = None
        ultimo_t_manita = None

    try:
        while True:
            try:
                crudo = ser.readline()
            except serial.SerialException:
                # El puerto fallo de plano (tipico si estabas conectada
                # DIRECTO a PEPE por Bluetooth y apagaste la manita, o se
                # perdio la señal). Se cierra la medicion en curso y se
                # reintenta abrir el puerto solo -- el script sigue vivo,
                # esperando a que enciendas la manita de nuevo.
                print("[AVISO] Se perdio la conexion con el puerto serie -- cerrando la medicion en curso...")
                cerrar_medicion_actual()
                ser.close()
                ser = abrir_puerto(puerto)
                print("[INFO] Puerto reabierto. Esperando datos...")
                continue

            if not crudo:
                # Timeout sin datos: si ya hay una medicion en curso y
                # llevamos un buen rato sin ninguna lectura real, se
                # considera que la manita se apago -- se cierra sola.
                if (
                    filas
                    and ultima_lectura_real is not None
                    and (time.time() - ultima_lectura_real) > SIN_DATOS_CIERRA_MEDICION_S
                ):
                    print("[INFO] Sin señal de la manita por un rato -- cerrando esta medicion...")
                    cerrar_medicion_actual()
                continue

            linea = crudo.decode("utf-8", errors="replace").strip()
            if not linea:
                continue

            archivo_log.write(linea + "\n")
            archivo_log.flush()

            datos = parsear_linea(linea)

            if not datos:
                # Linea sin pares etiqueta:valor (ruido, o si usas la ESP32,
                # sus propios mensajes de estado): no cuenta como lectura
                # real, se ignora para la hoja de calculo.
                print(f"[..] {linea}")
                continue

            # "t" (reloj de la manita) no es columna: se usa solo para el tiempo.
            t_manita = datos.pop("t", None)

            # Solo se aceptan lineas COMPLETAS (las 13 variables). La primera
            # linea que llega al abrir el puerto casi siempre viene cortada a
            # la mitad (por ejemplo solo "D4:2500"); antes se colaba como
            # fila con celdas vacias y, como no traia "t", hacia que TODA la
            # medicion usara la hora de llegada en vez del reloj de la manita.
            faltan = [c for c in COLUMNAS_ORDEN if c not in datos]
            if faltan:
                print(f"[..] linea incompleta descartada: {linea}")
                continue
            if t_manita is None and t0_manita is not None:
                # Esta medicion usa el reloj de la manita; una linea sin "t"
                # es una linea danada en el camino: se descarta.
                print(f"[..] linea sin 't' descartada: {linea}")
                continue

            # Si el reloj de la manita "volvio atras", es que se reinicio
            # (la apagaste y prendiste rapido, sin llegar a los segundos de
            # silencio): eso tambien cuenta como una medicion nueva.
            if (
                t_manita is not None
                and ultimo_t_manita is not None
                and t_manita < ultimo_t_manita
                and filas
            ):
                print("[INFO] La manita se reinicio -- cerrando esta medicion...")
                cerrar_medicion_actual()

            ultima_lectura_real = time.time()

            if t0 is None:
                # Primera lectura real de esta medicion: arranca el
                # cronometro en 0 y define el archivo de esta medicion.
                t0 = ultima_lectura_real
                t0_manita = t_manita
                archivo_excel_actual, archivo_csv_actual = nuevo_nombre_archivo()
                print(f"[INFO] Nueva medicion iniciada -> {archivo_excel_actual.name}")

            if t_manita is not None and t0_manita is not None:
                # Tiempo REAL de la lectura, medido por la propia manita.
                segundos = round((t_manita - t0_manita) / 1000.0, 3)
            else:
                # Firmware viejo sin "t": hora de llegada a la computadora.
                segundos = round(ultima_lectura_real - t0, 3)
            ultimo_t_manita = t_manita

            fila = {"t_segundos": segundos, **datos}
            filas.append(fila)

            # A ~30 lecturas/seg imprimir cada una satura la terminal, asi
            # que solo se muestra 1 de cada 30 (mas o menos 1 vez por
            # segundo) como confirmacion de que sigue llegando informacion.
            # Todas las lecturas se siguen guardando igual en Excel/CSV.
            if len(filas) % 30 == 0:
                print(f"[OK] ({len(filas)} lecturas, t={segundos:.2f}s) {fila}")

            filas_desde_guardado += 1
            ahora = time.time()
            if (
                ahora - ultimo_guardado >= GUARDAR_EXCEL_CADA_S
                or filas_desde_guardado >= GUARDAR_EXCEL_CADA_N_FILAS
            ):
                guardar_excel_datos(columnas, filas, archivo_excel_actual)
                guardar_csv(columnas, filas, archivo_csv_actual)
                ultimo_guardado = ahora
                filas_desde_guardado = 0

    except KeyboardInterrupt:
        print("\n[INFO] Deteniendo captura...")
    finally:
        if filas:
            finalizar_medicion(columnas, filas, archivo_excel_actual, archivo_csv_actual)
        archivo_log.close()
        ser.close()
        print("[INFO] Listo.")


if __name__ == "__main__":
    main()
