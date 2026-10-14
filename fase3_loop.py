"""
Fase 3 - Loop continuo webcam -> Excel.
Abre camara y libro una sola vez, y en cada iteracion solo repinta las
celdas que han cambiado respecto al frame anterior.

Ejecutar:
  python fase3_loop.py

Parar: Ctrl+C en la terminal.
"""

import os
import time
import cv2
import numpy as np
import xlwings as xw

RES = 30
LEVELS = 6
INVERT = False
TARGET_FPS = 6  # limite superior; si el pintado es lento, ira mas despacio

HERE = os.path.dirname(os.path.abspath(__file__))
TEMPLATE = os.path.join(HERE, "camera.xlsx")

GRID_ROW0 = 1
GRID_COL0 = 3

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


def quantize(gray, levels, invert):
    q = (gray.astype(np.int32) * levels // 256)
    q = np.clip(q, 0, levels - 1)
    if invert:
        q = (levels - 1) - q
    return q


def ensure_template(app):
    if os.path.exists(TEMPLATE):
        return app.books.open(TEMPLATE)
    wb = app.books.add()
    sht = wb.sheets[0]
    sht.name = "Camara"
    for i, rgb in enumerate(DEFAULT_PALETTE[:LEVELS]):
        cell = sht.range((i + 1, 1))
        cell.value = f"Nivel {i}"
        cell.color = rgb
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


def paint_delta(sht, new_m, old_m, palette):
    """Pinta solo los runs de celdas cuyo valor ha cambiado."""
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
            # extender mientras: siguiente celda cambio Y mantiene el mismo valor nuevo
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
        print("ERROR: no se pudo abrir la webcam.")
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

        palette = read_palette(sht, LEVELS)
        frame_period = 1.0 / TARGET_FPS
        prev = None
        frame_idx = 0
        t_start = time.time()

        print(f"Loop iniciado ({RES}x{RES}, {LEVELS} niveles, {TARGET_FPS} fps obj). Ctrl+C para salir.")
        try:
            while True:
                t0 = time.time()
                gray = capture(cap, RES)
                if gray is None:
                    continue
                matrix = quantize(gray, LEVELS, INVERT)

                calls = paint_delta(sht, matrix, prev, palette)
                prev = matrix
                frame_idx += 1

                dt = time.time() - t0
                if frame_idx % 10 == 0:
                    fps_real = frame_idx / (time.time() - t_start)
                    print(f"frame {frame_idx}: {calls} runs, {dt*1000:.0f}ms, fps real {fps_real:.1f}")

                # cada ~2s, releer paleta por si el usuario cambio colores
                if frame_idx % (TARGET_FPS * 2) == 0:
                    try:
                        palette = read_palette(sht, LEVELS)
                        prev = None  # fuerza repintado completo
                    except Exception:
                        pass

                remaining = frame_period - dt
                if remaining > 0:
                    time.sleep(remaining)
        except KeyboardInterrupt:
            print("\nParando...")
    finally:
        try:
            cap.release()
        except Exception:
            pass
        print("Webcam liberada. Excel queda abierto.")


if __name__ == "__main__":
    main()
