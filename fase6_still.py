"""
Fase 6 - Captura unica en alta resolucion.
- UNA sola foto (sin loop).
- Celdas lo mas pequenas posibles.
- Color via subpixeles RGB puros como en fase5 (cada pixel = 3 celdas R/G/B).
- Mas niveles de cuantizacion porque el tiempo de pintado ya no limita a fps.

Ejecutar:
  python fase6_still.py

Ajustes rapidos:
  RES            -> resolucion (NxN pixeles)
  LEVELS         -> niveles por canal (8, 12, 16...)
  CELL_COL_WIDTH -> ancho de columna en Excel (baja -> celdas mas pequenas)
  CELL_ROW_HEIGHT-> alto de fila en Excel
"""

import os
import time
import cv2
import numpy as np
import xlwings as xw

RES = 600                 # imagen 600x600. Grid Excel: 600 filas x 1800 cols (~1.08 M celdas).
LEVELS = 12               # foto fija -> podemos subirlo sin coste de fps
SUBPX = 3                 # celdas por pixel (R, G, B)

# Cuanto mas bajos, mejor se fusionan las rayas R/G/B al mirar con zoom out.
CELL_COL_WIDTH = 0.5      # Excel admite ~0.08 minimo; 0.5 ~= 4-5 px en pantalla
CELL_ROW_HEIGHT = 4       # puntos (default es 15). 4 ~= 5-6 px.

HERE = os.path.dirname(os.path.abspath(__file__))
TEMPLATE_XLSM = os.path.join(HERE, "camera.xlsm")
TEMPLATE_XLSX = os.path.join(HERE, "camera.xlsx")
TEMPLATE = TEMPLATE_XLSM if os.path.exists(TEMPLATE_XLSM) else TEMPLATE_XLSX

GRID_ROW0 = 1
GRID_COL0 = 3

# stripe 0 (rojo) -> canal 2 de BGR; stripe 1 (verde) -> 1; stripe 2 (azul) -> 0
STRIPE_CHANNEL = [2, 1, 0]


def open_camera():
    """Abre la webcam pidiendo la maxima resolucion nativa."""
    for name, be in [("MSMF", cv2.CAP_MSMF), ("DSHOW", cv2.CAP_DSHOW), ("ANY", cv2.CAP_ANY)]:
        for idx in (0, 1, 2):
            cap = cv2.VideoCapture(idx, be)
            if cap.isOpened():
                cap.set(cv2.CAP_PROP_FRAME_WIDTH, 1920)
                cap.set(cv2.CAP_PROP_FRAME_HEIGHT, 1080)
                ok, _ = cap.read()
                if ok:
                    w = int(cap.get(cv2.CAP_PROP_FRAME_WIDTH))
                    h = int(cap.get(cv2.CAP_PROP_FRAME_HEIGHT))
                    print(f"Camara abierta: backend={name}, index={idx}, {w}x{h}")
                    return cap
                cap.release()
    return None


def capture_color(cap, res):
    # descartar frames iniciales para que la camara ajuste exposicion/balance
    for _ in range(8):
        cap.read()
    ok, frame = cap.read()
    if not ok:
        return None
    frame = cv2.flip(frame, 1)
    # recorte central cuadrado antes del downscale para no deformar la imagen
    h, w = frame.shape[:2]
    s = min(h, w)
    y0 = (h - s) // 2
    x0 = (w - s) // 2
    frame = frame[y0:y0 + s, x0:x0 + s]
    return cv2.resize(frame, (res, res), interpolation=cv2.INTER_AREA)


def quantize_color(bgr, levels):
    q = (bgr.astype(np.int32) * levels // 256)
    return np.clip(q, 0, levels - 1)


def intensity_for(level, levels):
    if levels <= 1:
        return 0
    return int(level * 255 / (levels - 1))


def set_screen_updating(app, value, attempts=10):
    """Excel puede estar ocupado (procesando el resize del grid) y rechazar
    SetScreenUpdating con 0x800ac472 (VBA_E_IGNORE). Reintentamos con espera.
    Si tras los intentos sigue sin aceptar, devolvemos False y seguimos sin
    desactivar ScreenUpdating (mas lento pero funcional).
    """
    for i in range(attempts):
        try:
            app.api.ScreenUpdating = value
            return True
        except Exception:
            time.sleep(0.3 * (i + 1))
    return False


def ensure_template(app):
    if os.path.exists(TEMPLATE):
        wb = app.books.open(TEMPLATE)
    else:
        wb = app.books.add()
        wb.sheets[0].name = "Camara"
        wb.save(TEMPLATE)

    sht = wb.sheets[0]

    # limpiar paleta / panel de control residual de fases anteriores
    sht.range((1, 1), (20, 2)).value = None
    sht.range((1, 1), (20, 2)).color = None

    # redimensionar grid
    grid = sht.range(
        (GRID_ROW0, GRID_COL0),
        (GRID_ROW0 + RES - 1, GRID_COL0 + RES * SUBPX - 1),
    )
    grid.column_width = CELL_COL_WIDTH
    grid.row_height = CELL_ROW_HEIGHT

    # limpiar zona residual amplia (por si venimos de un grid mas grande)
    sht.range(
        (GRID_ROW0, GRID_COL0),
        (GRID_ROW0 + max(RES, 40) - 1, GRID_COL0 + max(RES * SUBPX, 120) - 1),
    ).color = None

    return wb


def _set_color_retry(rng, rgb, attempts=5):
    """Reintenta asignar color si Excel esta momentaneamente ocupado."""
    for i in range(attempts):
        try:
            rng.color = rgb
            return
        except Exception:
            time.sleep(0.2 * (i + 1))
    rng.color = rgb  # ultimo intento sin capturar -> propaga si falla


def paint_full_color(sht, matrix, levels):
    """Pinta el grid completo. Fusiona runs verticales del mismo nivel en cada
    columna de cada stripe para minimizar llamadas COM.
    matrix: shape (rows, cols, 3) en BGR con valores 0..levels-1.
    """
    rows_n, cols_n, _ = matrix.shape
    total_cols = SUBPX * cols_n
    total = 0
    done_cols = 0
    t0 = time.time()
    progress_step = max(1, total_cols // 40)  # ~40 lineas de progreso
    for s in range(SUBPX):
        ch = STRIPE_CHANNEL[s]
        for c in range(cols_n):
            excel_col = GRID_COL0 + SUBPX * c + s
            col = matrix[:, c, ch]
            r = 0
            while r < rows_n:
                lvl = int(col[r])
                r_end = r
                while r_end + 1 < rows_n and int(col[r_end + 1]) == lvl:
                    r_end += 1
                intensity = intensity_for(lvl, levels)
                rgb = [0, 0, 0]
                rgb[s] = intensity  # s=0 rojo, s=1 verde, s=2 azul
                top = (GRID_ROW0 + r, excel_col)
                if r == r_end:
                    rng = sht.range(top)
                else:
                    rng = sht.range(top, (GRID_ROW0 + r_end, excel_col))
                _set_color_retry(rng, tuple(rgb))
                total += 1
                r = r_end + 1
            done_cols += 1
            if done_cols % progress_step == 0:
                elapsed = time.time() - t0
                pct = done_cols / total_cols * 100
                eta = elapsed / done_cols * (total_cols - done_cols)
                print(f"  {pct:5.1f}% - col {done_cols}/{total_cols}, "
                      f"{total} runs, {elapsed:.0f}s elapsed, ~{eta:.0f}s ETA")
    return total


def main():
    cap = open_camera()
    if cap is None:
        print("ERROR: webcam no disponible.")
        return

    app = xw.App(visible=True, add_book=False)
    app.display_alerts = False
    try:
        wb = ensure_template(app)
        sht = wb.sheets[0]
        try:
            wb.activate()
            sht.activate()
        except Exception:
            pass
        # el resize de un grid grande deja a Excel ocupado un rato;
        # damos margen antes de la siguiente llamada COM.
        # Con RES=600 el resize/limpiado de ~1M celdas necesita varios segundos.
        time.sleep(5.0)

        print(f"Capturando frame {RES}x{RES}...")
        bgr = capture_color(cap, RES)
        if bgr is None:
            print("ERROR: no se pudo capturar de la webcam.")
            return

        matrix = quantize_color(bgr, LEVELS)

        total_cells = RES * RES * SUBPX
        print(f"Pintando grid {RES}x{RES * SUBPX} = {total_cells} celdas "
              f"(LEVELS={LEVELS})...")
        t0 = time.time()
        screen_off = set_screen_updating(app, False)
        if not screen_off:
            print("(Aviso: Excel ocupado, pintado sin ScreenUpdating=False)")
        try:
            calls = paint_full_color(sht, matrix, LEVELS)
        finally:
            if screen_off:
                set_screen_updating(app, True)
        dt = time.time() - t0
        print(f"OK: {calls} runs COM en {dt:.1f}s "
              f"(compresion {calls / total_cells * 100:.1f}%).")

        wb.save()
    finally:
        try:
            cap.release()
        except Exception:
            pass
        print("Webcam liberada. Excel queda abierto.")


if __name__ == "__main__":
    main()
