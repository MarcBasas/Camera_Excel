"""
Fase 5 - Color via subpixeles RGB.
- Cada pixel de la camara se pinta como 3 celdas en fila: R, G, B puros.
- Al combinarlos visualmente (distancia / zoom out) se percibe color.
- Panel de control identico al de fase4 (A12:B16).
- La paleta A1:A10 se ignora en modo color (los 3 stripes son rojo/verde/azul puros,
  con intensidad proporcional al nivel cuantizado).

Ejecutar:
  python fase5_color.py
"""

import os
import time
import cv2
import numpy as np
import xlwings as xw

RES = 24                 # imagen 24x24. Grid Excel resultante: 24 filas x 72 cols.
MAX_LEVELS = 10
TARGET_FPS = 4           # ojo: 3x mas celdas que fase4, bajamos fps
SUBPX = 3                # celdas por pixel (R, G, B)

# Geometria de celda. Baja ambos para que las rayas R/G/B se fusionen mejor.
CELL_COL_WIDTH = 1       # ancho de columna (unidades xlwings). Antes: 2.
CELL_ROW_HEIGHT = 6      # alto de fila (puntos). Antes: 12.

HERE = os.path.dirname(os.path.abspath(__file__))
TEMPLATE_XLSM = os.path.join(HERE, "camera.xlsm")
TEMPLATE_XLSX = os.path.join(HERE, "camera.xlsx")
TEMPLATE = TEMPLATE_XLSM if os.path.exists(TEMPLATE_XLSM) else TEMPLATE_XLSX

GRID_ROW0 = 1
GRID_COL0 = 3

CTRL_ROW0 = 12
CTRL_LABELS = ["LEVELS", "INVERT", "CONTRAST", "BRILLO", "STATUS"]
CTRL_DEFAULTS = {"LEVELS": 4, "INVERT": 0, "CONTRAST": 1.0, "BRILLO": 0, "STATUS": "RUN"}

# Orden visual de los stripes: 0=rojo, 1=verde, 2=azul.
# OpenCV lee en BGR, asi que el canal que corresponde a cada stripe es:
#   stripe 0 (rojo)  -> canal 2 (R) de BGR
#   stripe 1 (verde) -> canal 1 (G) de BGR
#   stripe 2 (azul)  -> canal 0 (B) de BGR
STRIPE_CHANNEL = [2, 1, 0]


def col_letter(n):
    s = ""
    while n > 0:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


def open_camera():
    for name, be in [("MSMF", cv2.CAP_MSMF), ("DSHOW", cv2.CAP_DSHOW), ("ANY", cv2.CAP_ANY)]:
        for idx in (0, 1, 2):
            cap = cv2.VideoCapture(idx, be)
            if cap.isOpened():
                ok, _ = cap.read()
                if ok:
                    print(f"Camara abierta: backend={name}, index={idx}")
                    return cap
                cap.release()
    return None


def capture_color(cap, res):
    ok, frame = cap.read()
    if not ok:
        return None
    frame = cv2.flip(frame, 1)
    return cv2.resize(frame, (res, res), interpolation=cv2.INTER_AREA)


def quantize_color(bgr, levels, invert, contrast, brillo):
    adj = cv2.convertScaleAbs(bgr, alpha=contrast, beta=brillo)
    q = (adj.astype(np.int32) * levels // 256)
    q = np.clip(q, 0, levels - 1)
    if invert:
        q = (levels - 1) - q
    return q  # shape (res, res, 3) en BGR con valores 0..levels-1


def intensity_for(level, levels):
    if levels <= 1:
        return 0
    return int(level * 255 / (levels - 1))


def ensure_template(app):
    if os.path.exists(TEMPLATE):
        wb = app.books.open(TEMPLATE)
    else:
        wb = app.books.add()
        wb.sheets[0].name = "Camara"
        wb.save(TEMPLATE)

    sht = wb.sheets[0]

    # Limpiar paleta residual de fase4 (A1:A10 con "Nivel 0".."Nivel 9")
    palette_range = sht.range((1, 1), (10, 1))
    palette_range.value = None
    palette_range.color = None

    # Redimensionar grid a RES x (RES*SUBPX)
    grid = sht.range(
        (GRID_ROW0, GRID_COL0),
        (GRID_ROW0 + RES - 1, GRID_COL0 + RES * SUBPX - 1),
    )
    grid.column_width = CELL_COL_WIDTH
    grid.row_height = CELL_ROW_HEIGHT

    # Limpiar zona residual (por si venimos de fase4: grid 30x30 gris)
    sht.range(
        (GRID_ROW0, GRID_COL0),
        (GRID_ROW0 + 35, GRID_COL0 + max(RES * SUBPX, 35) - 1),
    ).color = None

    # Panel de control
    for i, label in enumerate(CTRL_LABELS):
        label_cell = sht.range((CTRL_ROW0 + i, 1))
        value_cell = sht.range((CTRL_ROW0 + i, 2))
        if label_cell.value != label:
            label_cell.value = label
        if value_cell.value is None:
            value_cell.value = CTRL_DEFAULTS[label]
    wb.save(TEMPLATE)
    return wb


def read_control(sht):
    raw = {}
    for i, label in enumerate(CTRL_LABELS):
        raw[label] = sht.range((CTRL_ROW0 + i, 2)).value

    def _int(v, d):
        try:
            return int(v) if v is not None else d
        except (TypeError, ValueError):
            return d

    def _float(v, d):
        try:
            return float(v) if v is not None else d
        except (TypeError, ValueError):
            return d

    levels = max(2, min(MAX_LEVELS, _int(raw["LEVELS"], 4)))
    invert = bool(_int(raw["INVERT"], 0))
    contrast = max(0.1, min(5.0, _float(raw["CONTRAST"], 1.0)))
    brillo = max(-100, min(100, _int(raw["BRILLO"], 0)))
    status = str(raw["STATUS"] or "RUN").strip().upper()
    return dict(LEVELS=levels, INVERT=invert, CONTRAST=contrast, BRILLO=brillo, STATUS=status)


def paint_delta_color(sht, new_m, old_m, levels):
    """Recorre cada stripe (R, G, B) columna a columna, fusionando runs verticales
    del mismo nivel para minimizar llamadas a Excel.
    new_m / old_m: shape (rows, cols, 3) BGR con valores 0..levels-1.
    """
    rows_n, cols_n, _ = new_m.shape
    total = 0
    for s in range(SUBPX):
        ch = STRIPE_CHANNEL[s]
        for c in range(cols_n):
            excel_col = GRID_COL0 + SUBPX * c + s
            col_new = new_m[:, c, ch]
            col_old = old_m[:, c, ch] if old_m is not None else None
            r = 0
            while r < rows_n:
                lvl = int(col_new[r])
                if col_old is not None and int(col_old[r]) == lvl:
                    r += 1
                    continue
                r_end = r
                while r_end + 1 < rows_n and int(col_new[r_end + 1]) == lvl and (
                    col_old is None or int(col_old[r_end + 1]) != lvl
                ):
                    r_end += 1
                intensity = intensity_for(lvl, levels)
                rgb = [0, 0, 0]
                rgb[s] = intensity  # s=0 rojo, s=1 verde, s=2 azul
                top = (GRID_ROW0 + r, excel_col)
                if r == r_end:
                    rng = sht.range(top)
                else:
                    rng = sht.range(top, (GRID_ROW0 + r_end, excel_col))
                rng.color = tuple(rgb)
                total += 1
                r = r_end + 1
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
            wb.activate(); sht.activate()
        except Exception:
            pass
        time.sleep(0.5)

        ctrl = read_control(sht)
        prev = None
        prev_levels = ctrl["LEVELS"]
        frame_period = 1.0 / TARGET_FPS
        frame_idx = 0
        t_start = time.time()

        print(f"RUN color LEVELS={ctrl['LEVELS']} INVERT={ctrl['INVERT']} "
              f"CONTRAST={ctrl['CONTRAST']} BRILLO={ctrl['BRILLO']}")
        print("Escribe STOP en B16 para parar, o Ctrl+C aqui.")

        try:
            while True:
                t0 = time.time()
                bgr = capture_color(cap, RES)
                if bgr is None:
                    continue
                matrix = quantize_color(bgr, ctrl["LEVELS"], ctrl["INVERT"],
                                        ctrl["CONTRAST"], ctrl["BRILLO"])
                app.screen_updating = False
                try:
                    calls = paint_delta_color(sht, matrix, prev, ctrl["LEVELS"])
                finally:
                    app.screen_updating = True
                prev = matrix
                frame_idx += 1

                dt = time.time() - t0
                if frame_idx % 5 == 0:
                    fps_real = frame_idx / (time.time() - t_start)
                    print(f"frame {frame_idx}: {calls} runs, {dt*1000:.0f}ms, fps {fps_real:.1f}")

                if frame_idx % max(1, TARGET_FPS) == 0:
                    try:
                        new_ctrl = read_control(sht)
                    except Exception:
                        new_ctrl = ctrl
                    if new_ctrl["STATUS"] == "STOP":
                        print("STOP recibido desde Excel.")
                        break
                    if new_ctrl["LEVELS"] != prev_levels:
                        prev = None
                        prev_levels = new_ctrl["LEVELS"]
                    ctrl = new_ctrl

                remaining = frame_period - dt
                if remaining > 0:
                    time.sleep(remaining)
        except KeyboardInterrupt:
            print("\nCtrl+C.")
    finally:
        try:
            cap.release()
        except Exception:
            pass
        try:
            sht.range((CTRL_ROW0 + 4, 2)).value = "RUN"
            wb.save()
        except Exception:
            pass
        print("Webcam liberada. Excel queda abierto.")


if __name__ == "__main__":
    main()
