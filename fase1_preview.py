"""
Fase 1 - Prueba de concepto: webcam -> matriz cuantizada por consola.
Sin Excel todavia. Sirve para validar resolucion, umbrales y fps.

Controles:
  q = salir
  + / - = subir o bajar umbral entre niveles
  i     = invertir
  n     = cambiar numero de niveles (2, 3, 4, 5)
"""

import cv2
import numpy as np
import time

RES = 60              # resolucion (RES x RES)
LEVELS = 6            # numero de niveles de cuantizacion
INVERT = False
TARGET_FPS = 12

CHARS = [" ", ".", "o", "O", "#", "@"]  # del mas claro al mas oscuro

def quantize(gray, levels, invert):
    # gray: uint8 0..255 -> int 0..levels-1
    q = (gray.astype(np.int32) * levels // 256)
    q = np.clip(q, 0, levels - 1)
    if invert:
        q = (levels - 1) - q
    return q

def render_console(matrix, levels):
    # mapea niveles a caracteres visibles (si hay mas niveles que CHARS, reescala)
    scale = (len(CHARS) - 1) / max(levels - 1, 1)
    lines = []
    for row in matrix:
        lines.append("".join(CHARS[int(round(v * scale))] for v in row))
    print("\n".join(lines))

def main():
    global LEVELS, INVERT

    cap = cv2.VideoCapture(0, cv2.CAP_DSHOW)  # CAP_DSHOW = mas rapido en Windows
    if not cap.isOpened():
        print("No se puede abrir la webcam.")
        return

    print(f"Webcam abierta. Res={RES}x{RES}, niveles={LEVELS}. Ctrl+C para salir.")
    frame_period = 1.0 / TARGET_FPS

    try:
        while True:
            t0 = time.time()
            ok, frame = cap.read()
            if not ok:
                continue

            # espejo horizontal (mas intuitivo) y a gris
            frame = cv2.flip(frame, 1)
            gray = cv2.cvtColor(frame, cv2.COLOR_BGR2GRAY)

            # reducir con INTER_AREA (mejor para downscale)
            small = cv2.resize(gray, (RES, RES), interpolation=cv2.INTER_AREA)

            # cuantizar
            matrix = quantize(small, LEVELS, INVERT)

            # limpiar consola (Windows) y pintar
            print("\033[H\033[J", end="")  # reset cursor + clear
            print(f"Niveles={LEVELS}  Invert={INVERT}  FPS_obj={TARGET_FPS}")
            render_console(matrix, LEVELS)

            # mantener cadencia
            dt = time.time() - t0
            if dt < frame_period:
                time.sleep(frame_period - dt)
    except KeyboardInterrupt:
        pass
    finally:
        cap.release()
        print("\nFin.")

if __name__ == "__main__":
    main()
