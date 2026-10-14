# Camera in Excel

Your webcam, drawn live inside an Excel sheet. Every camera pixel becomes one or more cells, and Python repaints their background colors many times per second.

The project grew in six steps. Each `faseN_*.py` script is a working version of that step.

| Script | What it does |
| --- | --- |
| `fase1_preview.py` | Webcam to a quantized grid in the console, no Excel yet. Used to tune resolution and thresholds. |
| `fase2_frame_to_excel.py` | Paints one frame into `camera.xlsx` using a color palette you can edit in cells `A1:A10`. |
| `fase3_loop.py` | Continuous loop. Opens the camera and the workbook once, and repaints only the cells that changed. |
| `fase4_loop.py` | Adds a control panel in cells `A12:B16`: levels, invert, contrast, brightness and a RUN/STOP switch. |
| `fase5_color.py` | Color. Each pixel is three cells, pure red, green and blue, and your eye mixes them at a distance. |
| `fase6_still.py` | One high-resolution photo (600x600 pixels, about 1 million cells) instead of a loop. |

## Requirements

- Windows with Microsoft Excel installed
- Python 3.10 or newer
- `pip install opencv-python numpy xlwings`

## Run

```bash
python fase5_color.py
```

Stop a loop with `Ctrl+C`, or type `STOP` in cell `B16` for phases 4 and 5. The `camera.xlsm` workbook is the template the scripts open; `camera.xlsx` is created on first run if the template is missing.

## Tuning

Resolution, number of levels and cell size are constants at the top of each script (`RES`, `LEVELS`, `CELL_COL_WIDTH`, `CELL_ROW_HEIGHT`).
