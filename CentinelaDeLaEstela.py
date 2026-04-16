import winreg
import ctypes
import time
import logging
import os
import sys
import keyboard
import atexit
import threading
import win32gui
import win32con
from datetime import datetime
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler

# ---------- CONSTANTES ----------
# Win32 API constants para modificar la estela del cursor y notificar cambios
SPI_SETMOUSETRAILS = 0x005D  # Enable / set estela del cursor
SPI_GETMOUSETRAILS = 0x005E  # Obtener valor actual de estela del cursor
SPIF_UPDATEINIFILE = 0x01  # Guarda la configuración en el registro para que persista
SPIF_SENDCHANGE = 0x02  # Notifica a las apps y al sistema que el parámetro ha cambiado
# Intervalos de tiempo
SLEEP_INTERVAL = 7200  # 2 horas
CHECK_INTERVAL = 10  # segundos
# Constantes para el bloque de estadísticas en el log
STATS_START = "--- STATS ---"
STATS_END = "--- END STATS ---"

# ---------- VARIABLES ----------
# logger global para registrar eventos y estadísticas
logger = logging.getLogger("EstelaCursor")
logger.setLevel(logging.INFO)  # Nivel INFO: registra información y alertas
file_handler = None  # Handler del archivo de log
# Variables para control de estado y estadísticas
current_month = None  # Mes actual del log
last_control_state = None  # Evita duplicados en el LOG
observer = None  # Watchdog observer para monitorizar el archivo de control
# Contadores para estadísticas mensuales
deactivated_count = 0
modified_count = 0
error_count = 0
# Control de ejecución para el bucle
running = True
paused = False

# ----------**** FUNCIONES ****----------
# ---------- BASE DIR ----------
# Determina la carpeta desde donde se ejecuta el script o el .exe
def base_dir():
    """Retorna la carpeta donde se ejecuta el script o .exe"""
    if getattr(sys, "frozen", False):
        # Ejecutable (.exe)
        return os.path.dirname(sys.executable)
    else:
        # Script (.py)
        return os.path.dirname(os.path.abspath(__file__))

# ---------- CONFIGURACIÓN DE ENTORNO ----------
BASE_DIR = base_dir()
LOG_DIR = os.path.join(BASE_DIR, "LOG")
os.makedirs(LOG_DIR, exist_ok=True)  # Crea la carpeta si no existe
CONTROL_FILE = os.path.join(LOG_DIR, "controlDeEstela.txt")
# El contenido esperado del archivo:
# "REANUDAR" → script activo
# "PAUSAR"  → script pausado
# "SALIR"    → script debe detenerse

def current_log_path():
    """Retorna la ruta completa del log actual basado en el mes actual"""
    return os.path.join(LOG_DIR, f"estela_cursor_{current_month}.log")

# ---------- OBTENER MES PASADO DEL LOG ----------
def get_last_logged_month():
    """
    Busca en BASE_DIR el último archivo de log con formato:
    estela_cursor_MM-YYYY.log
    y devuelve 'MM-YYYY' o None si no hay ninguno.
    """
    # Si no hay logs, retornamos None para iniciar con el mes actual
    try:
        # Listar archivos que coincidan con el patrón de log mensual
        logs = [
            f
            for f in os.listdir(LOG_DIR)
            if f.startswith("estela_cursor_") and f.endswith(".log")
        ]

        # Si no hay logs, retornamos None para iniciar con el mes actual
        if not logs:
            return None

        # Extrae MM-YYYY y ordena cronológicamente
        months = [f.replace("estela_cursor_", "").replace(".log", "") for f in logs]
        months.sort(key=lambda m: datetime.strptime(m, "%m-%Y"))
        return months[-1]

    except Exception as e:
        logger.error(f"❌Error detectando último mes de log: {e}")
        return None

def load_stats_from_log(log_path):
    """
    Lee el bloque --- STATS --- del log mensual si existe
    y devuelve los contadores persistidos.
    """
    stats = {"desactivaciones": 0, "modificaciones": 0, "errores": 0}

    # Si el archivo no existe, retornamos los contadores en 0
    if not os.path.exists(log_path):
        return stats

    try:
        # Abrimos el log y buscamos el bloque de estadísticas
        with open(log_path, "r", encoding="utf-8") as f:
            in_stats = False
            for line in f:
                line = line.strip()
                if line == STATS_START:
                    in_stats = True
                    continue
                if line == STATS_END:
                    break
                if in_stats and "=" in line:
                    key, value = line.split("=", 1)
                    if key in stats:
                        stats[key] = int(value)
    except Exception as e:
        logger.error(f"❌Error leyendo STATS del log: {e}")

    # retornamos el diccionario con los contadores cargados (o 0 si no se pudieron cargar)
    return stats

def write_stats_to_log(log_path):
    """
    Escribe (o reemplaza) el bloque --- STATS --- al final del log.
    Garantiza que exista una sola vez.
    """
    # Si el log no existe, lo crea con el bloque de estadísticas
    try:
        lines = []
        if os.path.exists(log_path):
            with open(log_path, "r", encoding="utf-8") as f:
                lines = f.readlines()

        # Eliminar bloque STATS previo si existe
        cleaned = []
        skip = False
        for line in lines:
            if line.strip() == STATS_START:
                skip = True
                continue
            if skip and line.strip() == STATS_END:
                skip = False
                continue
            if not skip:
                cleaned.append(line)

        # Añadir bloque STATS actualizado
        cleaned.append(f"{STATS_START}\n")
        cleaned.append(f"desactivaciones={deactivated_count}\n")
        cleaned.append(f"modificaciones={modified_count}\n")
        cleaned.append(f"errores={error_count}\n")
        cleaned.append(f"{STATS_END}\n")

        with open(log_path, "w", encoding="utf-8") as f:
            f.writelines(cleaned)

    except Exception as e:
        logger.error(f"❌Error escribiendo STATS en el log: {e}")

# ---------- ROTACION MENSUAL ----------
def setup_logger():
    """
    Configura el logger para guardar los logs en un archivo mensual.
    Si cambia el mes, cierra el log anterior y crea uno nuevo.
    """
    global current_month, file_handler
    global deactivated_count, modified_count, error_count

    month_str = datetime.now().strftime("%m-%Y")  # mes actual
    # Detectamos si el mes ha cambiado comparando con el mes actual registrado
    month_changed = current_month is not None and month_str != current_month

    # Si el proceso es nuevo, intentamos recuperar el último mes desde los logs
    if current_month is None:
        last_logged_month = get_last_logged_month()
        if last_logged_month:
            current_month = last_logged_month
        else:
            current_month = month_str  # Si no hay logs, iniciamos con el mes actual

    if month_str == current_month and file_handler:
        return  # No hay cambio de mes y el handler ya existe

    # Si el mes ha cambiado, cerramos el handler del mes anterior
    if file_handler:
        logger.removeHandler(file_handler)
        file_handler.close()

    # Configurar nuevo handler para el mes actual
    log_path = os.path.join(LOG_DIR, f"estela_cursor_{month_str}.log")
    # Si el mes ha cambiado, reiniciamos los contadores para el nuevo mes
    if month_changed:
        logger.info("📆 Cambio de mes detectado → reiniciando estadísticas")
        deactivated_count = 0
        modified_count = 0
        error_count = 0
    # Si el mes no ha cambiado, intentamos cargar las estadísticas previas del log
    else:
        stats = load_stats_from_log(log_path)
        deactivated_count = stats["desactivaciones"]
        modified_count = stats["modificaciones"]
        error_count = stats["errores"]

    # Determinar si el archivo de log es nuevo para escribir la cabecera
    is_new_log_file = not os.path.exists(log_path)

    # Configurar el logger para el nuevo archivo
    file_handler = logging.FileHandler(log_path, encoding="utf-8")
    formatter = logging.Formatter(
        "%(asctime)s | %(levelname)s | %(message)s", datefmt="%Y-%m-%d %H:%M:%S"
    )
    file_handler.setFormatter(formatter)
    logger.addHandler(file_handler)

    # ✅ Cabecera SOLO si el archivo es nuevo
    if is_new_log_file:
        logger.info(f"=== Centinela de la Estela 📝LOG del mes {month_str} ===")

    # Actualizamos mes actual
    current_month = month_str

# ---------- OBTENER ESTELA DEL SISTEMA ----------
def get_mouse_trails_system():
    """
    Obtiene el valor REAL de la estela desde el sistema (no solo registro)
    """
    value = ctypes.c_int()
    ctypes.windll.user32.SystemParametersInfoW(
        SPI_GETMOUSETRAILS, 0, ctypes.byref(value), 0
    )
    return str(value.value)

# ----------** FUNCION PRINCIPAL **----------
def activar_estela():
    """
    Comprueba el valor actual de MouseTrails en el registro.
    - Si está desactivado o con longitud incorrecta, lo corrige automáticamente.
    - Registra en el log alertas y estadísticas.
    """
    global deactivated_count, modified_count, error_count

    if not current_month:
        return  # Salir en caso de que current_month = None

    DESIRED_VALUE = "7"  # valor deseado (0->desactivado 7->valor máximo)

    try:
        # Abrir clave de registro del usuario acutal
        key = winreg.OpenKey(
            winreg.HKEY_CURRENT_USER,
            r"Control Panel\Mouse",
            0,
            winreg.KEY_READ | winreg.KEY_SET_VALUE,
        )

        # Leer valor actual de MouseTrails
        registry_value, _ = winreg.QueryValueEx(key, "MouseTrails")  # Obtenemos el valor del registro
        system_value = get_mouse_trails_system()  # Obtenemos el valor real del sistema

        # ---------- DETECCION Y CLASIFICACION ----------
        if system_value != DESIRED_VALUE:

            # Diferenciar si el cambio viene del registro o solo del sistema (runtime)
            if registry_value == DESIRED_VALUE:
                logger.warning("👻 Cambio SOLO en memoria detectado (no persistente)")

            # Caso 1: estela desactivada completamente
            if system_value == "0":
                logger.warning("🚨 ALERTA: Estela del ratón DESACTIVADA externamente")
                deactivated_count += 1  # Contador de desactivaciones
                write_stats_to_log(current_log_path())

            # Caso 2: estela activa pero con longitud distinta
            elif system_value.isdigit() and 1 <= int(system_value) <= 6:
                logger.warning(
                    f"⚠️ Aviso: Longitud de estela modificada "
                    f"(system={system_value}, reg={registry_value} → {DESIRED_VALUE})"
                )
                modified_count += 1  # Contador de modificaciones
                write_stats_to_log(current_log_path())

            # Caso 3: valor inesperado / corrupto
            else:
                logger.warning(
                    f"⚠️❓ Valor inesperado de MouseTrails detectado: {system_value}"
                )
                error_count += 1  # Contador de errores
                write_stats_to_log(current_log_path())

            # ---------- CORRECCION DE LA ESTELA ----------
            # 1️⃣ Modifica el registro de Windows directamente
            if registry_value != DESIRED_VALUE:
                winreg.SetValueEx(
                    key,  # La clave de registro abierta
                    "MouseTrails",  # El nombre del valor que queremos cambiar
                    0,  # Reservado, siempre 0
                    winreg.REG_SZ,  # Tipo de valor: cadena de texto
                    DESIRED_VALUE,  # Valor que queremos poner (por ejemplo "7")
                )

            # 2️⃣ Aplicar los cambios inmediatamente en el sistema
            # SystemParametersInfoW permite notificar a Windows que se ha cambiado un parámetro de usuario
            result = ctypes.windll.user32.SystemParametersInfoW(
                SPI_SETMOUSETRAILS,
                int(DESIRED_VALUE),
                None,
                SPIF_UPDATEINIFILE | SPIF_SENDCHANGE,
            )
            if not result:
                logger.error("❌Error aplicando cambios")

            # 3️⃣ Registrar en el log que hemos corregido la estela
            logger.info(f"🔁Estela restaurada automáticamente a {DESIRED_VALUE}")

        else:
            logger.info(
                f"👀Estela en estado correcto (system={system_value}, reg={registry_value})"
            )

        # Cerramos la clave de registro para liberar recursos
        winreg.CloseKey(key)

    # Captura de errores
    except Exception as e:
        logger.error(f"❌Error en activar_estela(): {e}")

# ---------- SHUTDOWN COORDINADO ----------
def request_shutdown(reason: str):
    """
    Cierre coordinado del script: escribe logs finales,
    forzar flush y cambia running a False
    """
    global running

    # Escribimos en el log la razón del cierre y las estadísticas finales
    try:
        setup_logger()  # asegura handler activo
        logger.info(f"🛑 Evento de cierre detectado: {reason}")

        if current_month:
            log_path = current_log_path()
            write_stats_to_log(log_path)

        # Forzar flush inmediato
        for h in logger.handlers:
            h.flush()

    except Exception:
        pass

    running = False

# ---------- WM_QUERYENDSESSION ----------
def shutdown_wnd_proc(hwnd, msg, wparam, lparam):
    """Procesa mensajes WM_QUERYENDSESSION de Windows para apagar/logoff"""
    # Solicitamos un cierre coordinado
    if msg == win32con.WM_QUERYENDSESSION:
        request_shutdown("🖥️ (apagado/logoff)")
        return True
    return win32gui.DefWindowProc(hwnd, msg, wparam, lparam)

def start_shutdown_listener():
    """Crea una ventana oculta para escuchar mensajes de apagado/logoff de Windows"""
    try:
        # Configuramos una clase de ventana para recibir mensajes del sistema
        wc = win32gui.WNDCLASS()
        wc.lpfnWndProc = shutdown_wnd_proc
        wc.lpszClassName = "EstelaCursorShutdownListener"
        wc.hInstance = win32gui.GetModuleHandle(None)
        win32gui.RegisterClass(wc)

        # Creamos una ventana oculta para recibir mensajes del sistema
        win32gui.CreateWindow(
            wc.lpszClassName, wc.lpszClassName, 0, 0, 0, 0, 0, 0, 0, wc.hInstance, None
        )

        # Iniciamos el loop de mensajes para que la ventana pueda recibir eventos
        win32gui.PumpMessages()
    except Exception as e:
        try:
            setup_logger()
            logger.error(f"❌Error iniciando listener de apagado: {e}")
        except Exception as e:
            pass

# ---------- FUNCIONES DE CONTROL POR ARCHIVO ----------
def apply_control_state(state: str):
    """
    Aplica un estado leído desde el archivo de control:
    - "SALIR" → detiene el script completamente
    - "PAUSAR" → pausa el bucle principal
    - "REANUDAR" → reanuda el bucle principal
    """
    global running, paused, last_control_state

    # Normalizamos el estado a mayúsculas para evitar problemas de formato
    state = state.upper()

    if state == last_control_state:
        return  # Evita duplicados

    # Actualizamos el último estado aplicado para evitar logs repetidos
    last_control_state = state

    # Cambiamos el estado del script según el valor leído
    if state == "SALIR":
        request_shutdown("📋Archivo de control: 🔚SALIR")

    elif state == "PAUSAR":
        if not paused:
            paused = True
            logger.info("📋Archivo de control: ⏸️PAUSAR")

    elif state == "REANUDAR":
        if paused:
            paused = False
            logger.info("📋Archivo de control: ▶️REANUDAR")

def read_control_file():
    """
    Lee el archivo de control (controlDeEstela.txt) y aplica el estado indicado.
    Solo actúa si el archivo existe. Captura errores de lectura.
    """
    if os.path.exists(CONTROL_FILE):
        try:
            # Abrimos el archivo y aplicamos el estado
            with open(CONTROL_FILE, "r", encoding="utf-8") as f:
                apply_control_state(f.read().strip())
        except Exception as e:
            logger.error(f"❌Error leyendo archivo de control: {e}")

def write_control_file(state):
    """
    Escribe el estado actual en el archivo de control (controlDeEstela.txt).
    Esto permite persistir el estado entre ejecuciones y para hotkeys.
    """
    try:
        # Escribimos el estado en el archivo para que sea leído por el watchdog
        with open(CONTROL_FILE, "w", encoding="utf-8") as f:
            f.write(state)
    except Exception as e:
        logger.error(f"❌Error escribiendo archivo de control: {e}")

# ---------- WATCHDOG ----------
class ControlFileHandler(FileSystemEventHandler):
    """
    Manejador de eventos de Watchdog para detectar cambios en el archivo de control.
    Cada vez que el archivo es modificado, se lee y aplica el estado correspondiente.
    """

    def on_modified(self, event):
        # Solo actuamos si se modificó el archivo de control que nos interesa
        if os.path.abspath(event.src_path) == os.path.abspath(CONTROL_FILE):
            read_control_file()

# ---------- FUNCIONES HOTKEYS ----------
def stop_script():
    """
    Función llamada por la hotkey definida para salir del script.
    Cambia el estado a SALIR y escribe el archivo de control.
    """
    request_shutdown("⌨️Hotkey 🔚SALIR")

def toggle_pause():
    """
    Función llamada por la hotkey definida para pausar o reanudar el script.
    Cambia el estado a PAUSA/REANUDAR y escribe el archivo de control.
    """
    global paused
    paused = not paused
    write_control_file("PAUSAR" if paused else "REANUDAR")
    logger.info("⌨️Hotkey ⏸️PAUSA" if paused else "⌨️Hotkey ▶️REANUDAR")

# ---------- LIMPIEZA DE RECURSOS ----------
def cleanup():
    """
    Cierra observer de Watchdog, guarda estadísticas en LOG y limpia handlers del logger al salir
    """
    global observer

    # Escribimos en el log que el script se está cerrando
    try:
        logger.info("🧹Script cerrado")
    except Exception:
        pass

    # Detenemos el observer de Watchdog si está activo
    if observer:
        observer.stop()  # Detener observador de cambios
        observer.join()  # Esperar a que el hilo termine

    # Guardar estadísticas finales en el log antes de cerrar
    if current_month:
        log_path = current_log_path()
        write_stats_to_log(log_path)

    # Cierra y elimina todos los handlers del logger para liberar recursos
    for handler in logger.handlers[:]:
        handler.close()
        logger.removeHandler(handler)

# Registrar la función de limpieza para que se ejecute automáticamente al cerrar el script
atexit.register(cleanup)  # Asegura limpieza automática al cerrar

# ---------- DETECTAR APAGADO/REINICIO ----------
def console_ctrl_handler(ctrl_type):
    """
    Maneja eventos de cierre de consola o apagado de Windows.
    ctrl_type:
    0 = Ctrl+C
    1 = Ctrl+Break
    2 = Cerrar consola (X en ventana o cerrar sesión)
    5 = Apagado / Reinicio de Windows
    Solicita un cierre seguro.
    """
    # Mapeamos el tipo de evento a una descripción para el log
    ctrl_map = {
        0: "Ctrl+C",
        1: "Ctrl+Break",
        2: "Cierre de consola / cerrar sesión",
        5: "Apagado o reinicio del sistema",
    }

    # Obtenemos la razón del evento o un mensaje genérico si no está mapeado
    reason = ctrl_map.get(ctrl_type, f"Evento desconocido ({ctrl_type})")

    # Solicitamos un cierre coordinado con la razón del evento para que se registre en el log
    request_shutdown(f"🖥️Consola: {reason}")
    return True

# Registramos el handler para eventos de control de consola (incluye apagado/reinicio)
ctypes.windll.kernel32.SetConsoleCtrlHandler(
    ctypes.WINFUNCTYPE(
        ctypes.c_bool,  # Tipo de evento (Ctrl+C, cierre de consola, apagado, etc.)
        ctypes.c_uint,
    )(
        console_ctrl_handler
    ),  # Función que maneja el evento
    True,  # Activar el handler
)

# ---------- LOOP ----------
if __name__ == "__main__":
    # Limpiar handlers previos
    for handler in logger.handlers[:]:
        logger.removeHandler(handler)
        handler.close()

    # Inicializamos Watchdog para vigilar cambios en el archivo de control
    observer = Observer()
    setup_logger()  # Configura el logger, incluyendo rotación mensual
    logger.info("🚀Script iniciado")

    # --- Reinicio seguro del archivo de control al arrancar ---
    write_control_file("REANUDAR")

    # --- Registrar hotkeys para control manual ---
    try:
        keyboard.add_hotkey("ctrl+alt+shift+q", stop_script)  # salir
        keyboard.add_hotkey("ctrl+alt+shift+p", toggle_pause)  # pausar/reanudar
    except Exception as e:
        logger.error(f"❌No se pudieron registrar las hotkeys: {e}")

    # Configuramos el observer para monitorizar el archivo de control
    observer.schedule(ControlFileHandler(), LOG_DIR, recursive=False)
    observer.start()  # Iniciamos el hilo de observación

    threading.Thread(target=start_shutdown_listener, daemon=True).start()

    # Bucle principal: ejecuta activar_estela() cada X tiempo si no está pausado
    try:
        while running:
            setup_logger()  # Asegura que el logger esté actualizado y rotando mensualmente
            if not paused:
                activar_estela()

            # Espera granular en lugar de sleep largo
            elapsed = 0
            while running and elapsed < SLEEP_INTERVAL:
                time.sleep(CHECK_INTERVAL)
                elapsed += CHECK_INTERVAL

    except Exception as e:
        logger.error(f"❌Error en el bucle principal: {e}")
