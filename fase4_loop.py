"""
Fase 4 - Loop con panel de control en celdas.
- Paleta:  A1:A10  (se usan las primeras LEVELS filas)
- Control: A12:B16
    A12 LEVELS    B12 (int, 2..10)
    A13 INVERT    B13 (0 o 1)
    A14 CONTRAST  B14 (float, p.ej. 1.0 normal, 1.5 mas contraste)
    A15 BRILLO    B15 (int, -100..100)
    A16 STATUS    B16 (RUN para correr, STOP para detener)

Ejecutar:
  python fase4_loop.py

Parar: escribe STOP en B16 (o Ctrl+C en terminal).
"""

import os
import time
import cv2
import numpy as np
import xlwings as xw

RES = 30
MAX_LEVELS = 10
TARGET_FPS = 6

HERE = os.path.dirname(os.path.abspath(__file__))
TEMPLATE_XLSM = os.path.join(HERE, "camera.xlsm")
TEMPLATE_XLSX = os.path.join(HERE, "camera.xlsx")
TEMPLATE = TEMPLATE_XLSM if os.path.exists(TEMPLATE_XLSM) else TEMPLATE_XLSX

GRID_ROW0 = 1
GRID_COL0 = 3

# 10 grises por defecto (se reparten si el usuario no pone colores)
DEFAULT_PALETTE = [
    (255, 255, 255), (230, 230, 230), (205, 205, 205), (175, 175, 175),
    (145, 145, 145), (115, 115, 115), ( 85,  85,  85), ( 55,  55,  55),
    ( 25,  25,  25), (  0,   0,   0),
]

CTRL_ROW0 = 12  # fila inicial del panel de control
CTRL_LABELS = ["LEVELS", "INVERT", "CONTRAST", "BRILLO", "STATUS"]
CTRL_DEFAULTS = {
    "LEVELS": 6,
    "INVERT": 0,
    "CONTRAST": 1.0,
    "BRILLO": 0,
    "STATUS": "RUN",
}


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


def capture(cap, res):
    ok, frame = cap.read()
    if not ok:
        return None
    frame = cv2.flip(frame, 1)
    gray = cv2.cvtColor(frame, cv2.COLOR_BGR2GRAY)
    return cv2.resize(gray, (res, res), interpolation=cv2.INTER_AREA)


def quantize(gray, levels, invert, contrast, brillo):
    # contraste / brillo: out = alpha*in + beta
    g = cv2.convertScaleAbs(gray, alpha=contrast, beta=brillo)
    q = (g.astype(np.int32) * levels // 256)
    q = np.clip(q, 0, levels - 1)
    if invert:
        q = (levels - 1) - q
    return q


def ensure_template(app):
    """Crea o abre camera.xlsx y garantiza que existe el panel de control."""
    if os.path.exists(TEMPLATE):
        wb = app.books.open(TEMPLATE)
    else:
        wb = app.books.add()
        wb.sheets[0].name = "Camara"
        sht = wb.sheets[0]
        # Paleta por defecto en A1:A10
        for i, rgb in enumerate(DEFAULT_PALETTE):
            cell = sht.range((i + 1, 1))
            cell.value = f"Nivel {i}"
            cell.color = rgb
        # Grid
        grid = sht.range(
            (GRID_ROW0, GRID_COL0),
            (GRID_ROW0 + RES - 1, GRID_COL0 + RES - 1),
        )
        grid.column_width = 2
        grid.row_height = 12
        wb.save(TEMPLATE)

    sht = wb.sheets[0]
    # Asegurar panel de control: escribir solo si faltan labels o valores
    for i, label in enumerate(CTRL_LABELS):
        label_cell = sht.range((CTRL_ROW0 + i, 1))
        value_cell = sht.range((CTRL_ROW0 + i, 2))
        if label_cell.value != label:
            label_cell.value = label
        if value_cell.value is None:
            value_cell.value = CTRL_DEFAULTS[label]
    wb.save(TEMPLATE)
    return wb


def read_palette(sht, levels):
    colors = []
    for i in range(levels):
        c = sht.range((i + 1, 1)).color
        if c is None:
            # fallback: interpolar gris uniforme
            c = DEFAULT_PALETTE[min(i, len(DEFAULT_PALETTE) - 1)]
        colors.append(c)
    return colors


def read_control(sht):
    """Lee el panel y devuelve dict con tipos correctos."""
    raw = {}
    for i, label in enumerate(CTRL_LABELS):
        raw[label] = sht.range((CTRL_ROW0 + i, 2)).value

    def _int(v, default):
        try:
            return int(v) if v is not None else default
        except (TypeError, ValueError):
            return default

    def _float(v, default):
        try:
            return float(v) if v is not None else default
        except (TypeError, ValueError):
            return default

    levels = max(2, min(MAX_LEVELS, _int(raw["LEVELS"], 6)))
    invert = bool(_int(raw["INVERT"], 0))
    contrast = max(0.1, min(5.0, _float(raw["CONTRAST"], 1.0)))
    brillo = max(-100, min(100, _int(raw["BRILLO"], 0)))
    status = str(raw["STATUS"] or "RUN").strip().upper()
    return dict(LEVELS=levels, INVERT=invert, CONTRAST=contrast, BRILLO=brillo, STATUS=status)


def paint_delta(sht, new_m, old_m, palette):
    rows_n, cols_n = new_m.shape
    total = 0
    for ri in range(rows_n):
        row_new = new_m[ri]
        row_old = old_m[ri] if old_m is not None else None
        r = GRID_ROW0 + ri
        c = 0
        while c < cols_n:
            if row_old is not None and int(row_new[c]) == int(row_old[c]):
                c += 1
                continue
            val = int(row_new[c])
            c_end = c
            while c_end + 1 < cols_n and int(row_new[c_end + 1]) == val and (
                row_old is None or int(row_old[c_end + 1]) != val
            ):
                c_end += 1
            a1 = f"{col_letter(GRID_COL0 + c)}{r}"
            addr = a1 if c == c_end else f"{a1}:{col_letter(GRID_COL0 + c_end)}{r}"
            sht.range(addr).color = palette[val]
            total += 1
            c = c_end + 1
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
        palette = read_palette(sht, ctrl["LEVELS"])
        prev = None
        prev_levels = ctrl["LEVELS"]
        frame_period = 1.0 / TARGET_FPS
        frame_idx = 0
        t_start = time.time()

        print(f"RUN con LEVELS={ctrl['LEVELS']} INVERT={ctrl['INVERT']} "
              f"CONTRAST={ctrl['CONTRAST']} BRILLO={ctrl['BRILLO']}")
        print("Escribe STOP en B16 para parar, o Ctrl+C aqui.")

        try:
            while True:
                t0 = time.time()
                gray = capture(cap, RES)
                if gray is None:
                    continue
                matrix = quantize(gray, ctrl["LEVELS"], ctrl["INVERT"],
                                  ctrl["CONTRAST"], ctrl["BRILLO"])

                calls = paint_delta(sht, matrix, prev, palette)
                prev = matrix
                frame_idx += 1

                dt = time.time() - t0
                if frame_idx % 10 == 0:
                    fps_real = frame_idx / (time.time() - t_start)
                    print(f"frame {frame_idx}: {calls} runs, {dt*1000:.0f}ms, fps {fps_real:.1f}")

                # cada ~1s relee el panel de control
                if frame_idx % max(1, TARGET_FPS) == 0:
                    try:
                        new_ctrl = read_control(sht)
                    except Exception:
                        new_ctrl = ctrl
                    if new_ctrl["STATUS"] == "STOP":
                        print("STOP recibido desde Excel.")
                        break
                    # si cambio LEVELS o la paleta, recargar y forzar repintado total
                    if new_ctrl["LEVELS"] != prev_levels:
                        palette = read_palette(sht, new_ctrl["LEVELS"])
                        prev = None
                        prev_levels = new_ctrl["LEVELS"]
                    elif frame_idx % (TARGET_FPS * 3) == 0:
                        # cada ~3s refresca la paleta por si cambio algun color
                        palette = read_palette(sht, new_ctrl["LEVELS"])
                        prev = None
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
        # devolver STATUS a RUN para la proxima vez
        try:
            sht.range((CTRL_ROW0 + 4, 2)).value = "RUN"
            wb.save()
        except Exception:
            pass
        print("Webcam liberada. Excel queda abierto.")


if __name__ == "__main__":
    main()
