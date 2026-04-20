# Flow-Chart-Clicker

A visual, node-based automation tool for games and repetitive tasks. Build flowcharts of detection steps (image matching, color detection, OCR) and action steps (clicks, key presses, typed text) — no coding required.

## Features

- **PNG detection** — template-match images on screen with configurable threshold and image modes (Grayscale, Color, Binary)
- **Color detection** — find colors by HSV or RGB with tolerance, including pixel-exact and area-count modes
- **OCR number reading** — read numbers from the screen using Tesseract
- **Movement detection** — detect screen changes between frames
- **Human-like input** — configurable mouse speed, click variance, hold duration variance
- **GE Interface** — fetch live RuneScape Grand Exchange prices and inject them into typed actions
- **JSON import/export** — save and load flowcharts
- **PyInstaller-ready** — portable Tesseract support for packaged `.exe` distribution

## Setup

### 1. Install Python dependencies

```
pip install -r requirements.txt
```

### 2. Install Tesseract OCR (optional — only needed for Number/OCR steps)

**Option A — System install:** Download from https://github.com/UB-Mannheim/tesseract/wiki and install to `C:\Program Files\Tesseract-OCR\`

**Option B — Portable:** Place Tesseract-OCR contents in a `tesseract/` folder next to the script or `.exe`. The app will detect it automatically.

### 3. Run

```
python FlowchartClickerApp.py
```

## Hotkeys

| Key | Action |
|-----|--------|
| F2 | Start / Stop automation |
| F3 | Pick color or capture location |
| F4 | Select area |
| Ctrl+C | Copy selected steps |
| Ctrl+V | Paste steps |
| Delete | Delete selected steps |

## Building a distributable `.exe`

```
pip install pyinstaller
pyinstaller --onefile --windowed --icon=Flow2.ico FlowchartClickerApp.py
```

Place a `tesseract/` folder next to the resulting `.exe` for portable OCR support.

## Project structure

```
FlowchartClickerApp.py   # Entry point and main class
app/
  theme.py               # Dark theme + widget styling
  canvas.py              # Flowchart canvas drawing and mouse events
  panels.py              # UI panel builders (globals, log, testing, GE)
  properties.py          # Properties panel and step editing
  executor.py            # Automation execution engine
  detection.py           # Image/color/OCR detection algorithms
  mouse_actions.py       # Mouse movement and click execution
  ge.py                  # Grand Exchange API and price logic
  capture.py             # Screen capture, area selection, snipping
  fileops.py             # JSON I/O, step management, clipboard
  overlays.py            # Area overlay windows
  utils.py               # Logging, hotkeys, miscellaneous
```

## Dependencies

- `opencv-python` — image template matching and color detection
- `numpy` — array operations for CV algorithms
- `pyautogui` — mouse/keyboard control and screenshots
- `keyboard` — global hotkey binding
- `Pillow` — image handling and display
- `pytesseract` — OCR interface (requires Tesseract binary)

Creator: u/WonRu2 | License: Freeware
