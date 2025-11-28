import json
import time
import csv
import win32api
import win32con
import win32clipboard
import sys
from pathlib import Path

# Comando de cmd para crear el ejecutable: pyinstaller --onefile --icon=audib.ico "drag_cursor_v3,2.py"
# ..... O para la versión sin cmd visible: pyinstaller --onefile --noconsole --icon=audib.ico "drag_cursor_v3,2.py"
# ..... Para invocar un "no console" sin que truene: crear archivo run_silent.vbs: CreateObject("Wscript.Shell").Run "drag_cursor_v3,1.py", 0, True

SCRIPT_DIR = (
    Path(sys.executable).resolve().parent
    if getattr(sys, "frozen", False)
    else Path(__file__).resolve().parent
)

BASE_DIR = None

# Carga config.json desde la misma carpeta del script/exe.
def load_config(name):
    global BASE_DIR
    path = SCRIPT_DIR / name
    BASE_DIR = path.parent

    with open(path, "r", encoding="utf-8") as f:
        return json.load(f)


# Convierte rutas relativas del JSON a rutas absolutas seguras.
def resolver_ruta(ruta_str):
    ruta = Path(ruta_str)
    if not ruta.is_absolute():
        ruta = (BASE_DIR / ruta).resolve()
    return ruta

# CSV con UTF-8
def guardar_csv(texto, ruta_csv):
    ruta_csv = resolver_ruta(ruta_csv)

    ruta_csv.parent.mkdir(parents=True, exist_ok=True)

    #texto = reparar_utf8(texto) # ...... Añadí esto pensando que los datos se copiaban con error, pero luego descbrí que excel muestra los csv con un formato raro... en fin, al parecer no es necesario, pero no lo quiero borrar
    lineas = texto.splitlines()

    with open(ruta_csv, "w", encoding="utf-8", newline="") as f:
        writer = csv.writer(f)
        for linea in lineas:
            writer.writerow(linea.split("\t"))

    print(f"[+] CSV guardado: {ruta_csv}")


# DELAYS

MICRO_DELAY = 0.05
COPY_DELAY = 0.25


def select_area_shift_click(x1, y1, x2, y2):
    # 1) Click inicial
    win32api.SetCursorPos((x1, y1))
    time.sleep(MICRO_DELAY)
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN, 0, 0)
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP, 0, 0)
    time.sleep(MICRO_DELAY)

    # 2) Mantener Shift
    win32api.keybd_event(win32con.VK_SHIFT, 0, 0, 0)
    time.sleep(MICRO_DELAY)

    # 3) Click final mientras Shift está presionado
    win32api.SetCursorPos((x2, y2))
    time.sleep(MICRO_DELAY)
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN, 0, 0)
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP, 0, 0)
    time.sleep(MICRO_DELAY)

    # 4) Soltar Shift
    win32api.keybd_event(win32con.VK_SHIFT, 0, win32con.KEYEVENTF_KEYUP, 0)
    time.sleep(MICRO_DELAY)



def perform_copy():
    win32api.keybd_event(win32con.VK_CONTROL, 0, 0, 0)
    win32api.keybd_event(ord('C'), 0, 0, 0)

    time.sleep(MICRO_DELAY)

    win32api.keybd_event(ord('C'), 0, win32con.KEYEVENTF_KEYUP, 0)
    win32api.keybd_event(win32con.VK_CONTROL, 0, win32con.KEYEVENTF_KEYUP, 0)

    time.sleep(COPY_DELAY)


def copy_clipboard_unicode():
    win32clipboard.OpenClipboard()

    try:
        if win32clipboard.IsClipboardFormatAvailable(win32con.CF_UNICODETEXT):
            raw = win32clipboard.GetClipboardData(win32con.CF_UNICODETEXT)
        else:
            raw = ""
    finally:
        win32clipboard.CloseClipboard()

    #  ================= DECODIFICACIÓN UNIVERSAL =================
    # 1) Convertirlo a bytes
    try:
        b = raw.encode("latin1", errors="surrogateescape")
    except:
        return raw

    # 2) Intentar decodificar como UTF-8 real
    try:
        fixed = b.decode("utf-8")
        return fixed
    except:
        pass

    # 3) Si falla, regresamos algo legible sin fallar
    return raw


def reparar_utf8(texto):
    try:
        return texto.encode("latin1").decode("utf-8")
    except:
        return texto


def main():
    config = load_config("config.json")
    operaciones = config["operaciones"]

    for op in operaciones:
        x1 = op["x1"]
        y1 = op["y1"]
        x2 = op["x2"]
        y2 = op["y2"]

        csv_destino = op["csv"] 

        print(f"\n[•] Selección → {csv_destino}")

        #select_area(x1, y1, x2, y2)
        select_area_shift_click(x1, y1, x2, y2)

        perform_copy()

        texto = copy_clipboard_unicode()
        guardar_csv(texto, csv_destino)

    print("\n[+] Todo completado correctamente.")


if __name__ == "__main__":
    main()
