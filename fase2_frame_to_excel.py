"""
Fase 2 - Volcar UN frame a Excel con paleta de colores.
Crea camera.xlsx la primera vez con:
  - Paleta en A1..A{LEVELS}  (colores de fondo que el usuario puede editar)
  - Grid 60x60 a partir de C1 con celdas casi cuadradas

Ejecutar:
  python fase2_frame_to_excel.py

La primera vez crea el archivo y pinta un frame. Las siguientes lo abre, lee la
paleta actual (por si la has cambiado) y repinta con un frame nuevo.
"""

import os
import time
import cv2
import numpy as np
import xlwings as xw

RES = 30
LEVELS = 6
INVERT = False

HERE = os.path.dirname(os.path.abspath(__file__))
TEMPLATE = os.path.join(HERE, "camera.xlsx")

# posicion del grid en la hoja (1-based). C1 = fila 1, columna 3.
GRID_ROW0 = 1
GRID_COL0 = 3

# paleta por defecto: grises de claro a oscuro (RGB)
DEFAULT_PALETTE = [
    (255, 255, 255),
    (210, 210, 210),
    (165, 165, 165),
    (120, 120, 120),
    ( 70,  70,  70),
    (  0,   0,   0),
]


def col_letter(n):
    s = ""
    while n > 0:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


def row_runs(cols_sorted):
    """Lista de (c1, c2) con runs de columnas consecutivas."""
    runs = []
    start = prev = cols_sorted[0]
    for c in cols_sorted[1:]:
        if c == prev + 1:
            prev = c
        else:
            runs.append((start, prev))
            start = prev = c
    runs.append((start, prev))
    return runs


def ensure_template(app):
    if os.path.exists(TEMPLATE):
        return app.books.open(TEMPLATE)

    wb = app.books.add()
    sht = wb.sheets[0]
    sht.name = "Camara"

    # Paleta: etiqueta en columna A, color de fondo
    for i, rgb in enumerate(DEFAULT_PALETTE[:LEVELS]):
        cell = sht.range((i + 1, 1))
        cell.value = f"Nivel {i}"
        cell.color = rgb

    # Grid con celdas casi cuadradas
    grid = sht.range(
        (GRID_ROW0, GRID_COL0),
        (GRID_ROW0 + RES - 1, GRID_COL0 + RES - 1),
    )
    grid.column_width = 2
    grid.row_height = 12

    wb.save(TEMPLATE)
    return wb


def read_palette(sht, levels):
    colors = []
    for i in range(levels):
        c = sht.range((i + 1, 1)).color
        colors.append(c if c is not None else DEFAULT_PALETTE[i])
    return colors


def open_camera():
    """Prueba varios backends e indices hasta encontrar uno que funcione."""
    backends = [
        ("MSMF", cv2.CAP_MSMF),
        ("DSHOW", cv2.CAP_DSHOW),
        ("ANY", cv2.CAP_ANY),
    ]
    for name, be in backends:
        for idx in (0, 1, 2):
            cap = cv2.VideoCapture(idx, be)
            if cap.isOpened():
                ok, _ = cap.read()
                if ok:
                    print(f"Camara abierta: backend={name}, index={idx}")
                    return cap
                cap.release()
    return None


def capture_frame(res):
    cap = open_camera()
    if cap is None:
        return None
    # descartar unos frames para que la camara ajuste exposicion
    for _ in range(5):
        cap.read()
    ok, frame = cap.read()
    cap.release()
    if not ok:
        return None
    frame = cv2.flip(frame, 1)
    gray = cv2.cvtColor(frame, cv2.COLOR_BGR2GRAY)
    return cv2.resize(gray, (res, res), interpolation=cv2.INTER_AREA)


def quantize(gray, levels, invert):
    q = (gray.astype(np.int32) * levels // 256)
    q = np.clip(q, 0, levels - 1)
    if invert:
        q = (levels - 1) - q
    return q


def paint(sht, matrix, palette):
    """Pinta fila a fila por runs de mismo nivel. Cada llamada es un rango simple."""
    total_calls = 0
    rows_n, cols_n = matrix.shape
    for ri in range(rows_n):
        row = matrix[ri]
        r = GRID_ROW0 + ri
        c_start = 0
        while c_start < cols_n:
            val = int(row[c_start])
            c_end = c_start
            while c_end + 1 < cols_n and int(row[c_end + 1]) == val:
                c_end += 1
            a1 = f"{col_letter(GRID_COL0 + c_start)}{r}"
            addr = a1 if c_start == c_end else f"{a1}:{col_letter(GRID_COL0 + c_end)}{r}"
            sht.range(addr).color = palette[val]
            total_calls += 1
            c_start = c_end + 1
    return total_calls


def main():
    app = xw.App(visible=True, add_book=False)
    app.display_alerts = False
    try:
        wb = ensure_template(app)
        sht = wb.sheets[0]

        # activar la ventana y dar un pequeno margen por si Excel muestra
        # "Vista protegida" u otro dialogo que bloquee COM
        try:
            wb.activate()
            sht.activate()
        except Exception:
            pass
        time.sleep(0.5)

        palette = read_palette(sht, LEVELS)

        print("Capturando frame...")
        frame = capture_frame(RES)
        if frame is None:
            print("ERROR: no se pudo capturar de la webcam.")
            return
        matrix = quantize(frame, LEVELS, INVERT)

        print("Pintando en Excel...")
        t0 = time.time()
        screen_off = False
        try:
            app.api.ScreenUpdating = False
            screen_off = True
        except Exception as e:
            print(f"(Aviso: no se pudo desactivar ScreenUpdating: {e})")
        try:
            calls = paint(sht, matrix, palette)
        finally:
            if screen_off:
                try:
                    app.api.ScreenUpdating = True
                except Exception:
                    pass
        dt = time.time() - t0

        wb.save()
        print(f"OK: {calls} llamadas COM en {dt:.2f}s. Frame volcado.")
    finally:
        # dejamos Excel abierto para que el usuario lo inspeccione
        pass


if __name__ == "__main__":
    main()
