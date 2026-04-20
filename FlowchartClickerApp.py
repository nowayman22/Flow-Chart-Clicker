# -*- coding: utf-8 -*-
# --- Version 66 ---

import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext, ttk, colorchooser
import tkinter.font as tkfont
import os
import cv2
import numpy as np
import pyautogui
import time
import keyboard
import json
import random
import math
import copy
from PIL import Image, ImageTk # Visual Rep
import urllib.request # For API requests
import sys # For portable Tesseract path
import threading
import ctypes

# --- Helper for Portable Tesseract ---
def get_base_path():
    """ Get absolute path to resource, works for dev and for PyInstaller """
    # PyInstaller creates a temp folder and stores path in _MEIPASS
    if getattr(sys, 'frozen', False) and hasattr(sys, '_MEIPASS'):
        return sys._MEIPASS
    else:
        return os.path.abspath(".")

# --- OCR Dependency Check ---
try:
    import pytesseract
    
    # --- Path logic for portable Tesseract ---
    # The user should create a folder named 'tesseract' next to the script/EXE
    # and place the Tesseract-OCR contents there.
    portable_tesseract_path = os.path.join(get_base_path(), 'tesseract', 'tesseract.exe')
    
    if os.path.exists(portable_tesseract_path):
        pytesseract.pytesseract.tesseract_cmd = portable_tesseract_path
    else:
        # Fallback to the default installed location for development
        pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'
    
    pytesseract.get_tesseract_version() 
    PYTESSERACT_AVAILABLE = True

except (ImportError, FileNotFoundError, pytesseract.TesseractNotFoundError):
    PYTESSERACT_AVAILABLE = False


class FlowchartClickerApp:
    """
    An intelligent automation tool with an enhanced visual, node-based flowchart interface.
    """
    MULTIPLE_VALUES = "< multiple values >"

from app.theme import ThemeMixin
from app.canvas import CanvasMixin
from app.panels import PanelsMixin
from app.properties import PropertiesMixin
from app.executor import ExecutorMixin
from app.detection import DetectionMixin
from app.mouse_actions import MouseActionsMixin
from app.ge import GEMixin
from app.capture import CaptureMixin
from app.fileops import FileOpsMixin
from app.overlays import OverlaysMixin
from app.utils import UtilsMixin

__version__ = "1.0.0"


class FlowchartClickerApp(
    ThemeMixin, CanvasMixin, PanelsMixin, PropertiesMixin,
    ExecutorMixin, DetectionMixin, MouseActionsMixin, GEMixin,
    CaptureMixin, FileOpsMixin, OverlaysMixin, UtilsMixin
):
    """Visual node-based automation tool � see README.md for usage."""
    MULTIPLE_VALUES = "< multiple values >"

    def __init__(self, root):
        self.root = root
        self.root.title("Flowchart Automation Tool")
        try:
            # This works on Windows and some Linux DEs. Fails gracefully on others.
            self.root.iconbitmap("Flow2.ico")
        except tk.TclError:
            # self.log is not available yet, so we print
            print("Note: .ico format not supported on this system. Skipping icon.")

        screen_w = self.root.winfo_screenwidth()
        screen_h = self.root.winfo_screenheight()
        
        # Set window size to 90% of the screen
        win_w = int(screen_w * 0.9)
        win_h = int(screen_h * 0.9)

        # Calculate position to center the window
        pos_x = (screen_w // 2) - (win_w // 2)
        pos_y = (screen_h // 2) - (win_h // 2)

        self.root.geometry(f"{win_w}x{win_h}+{pos_x}+{pos_y}")
        self.root.minsize(1200, 700)
                
        # --- Theme Settings ---
        self.current_theme = {}
        self.status_label_color_state = 'blue'

        # --- Global Settings ---
        self.mouse_move_mode = tk.StringVar(value='Regular')
        self.mouse_speed = tk.DoubleVar(value=0.25)
        self.pixels_per_second = tk.IntVar(value=1000)
        self.min_move_time = tk.DoubleVar(value=0.05)
        self.max_move_time = tk.DoubleVar(value=0.3)
        self.scan_interval = tk.DoubleVar(value=0.25)
        self.hold_duration = tk.DoubleVar(value=0.08)
        self.loc_offset_variance = tk.IntVar(value=4)
        self.speed_variance = tk.DoubleVar(value=0.06)
        self.hold_duration_variance = tk.DoubleVar(value=0.03)
        self.hide_on_select = tk.BooleanVar(value=True)
        self.enable_all_show_area = tk.BooleanVar(value=False)
        self.start_at_stopped_pos = tk.BooleanVar(value=False)
        self.enable_dynamic_speed = tk.BooleanVar(value=False)

        # --- Flowchart Grid Settings ---
        self.grid_visible = tk.BooleanVar(value=False)
        self.grid_latching = tk.BooleanVar(value=False)
        self.grid_spacing = tk.IntVar(value=30)
        self.grid_opacity = tk.DoubleVar(value=0.3)

        self.area_x1 = tk.IntVar(value=0); self.area_y1 = tk.IntVar(value=0)
        self.area_x2 = tk.IntVar(value=screen_w); self.area_y2 = tk.IntVar(value=screen_h)
        
        self.global_settings_map = {
            'mouse_move_mode': {'model': self.mouse_move_mode, 'type': str},
            'mouse_speed': {'model': self.mouse_speed, 'type': float},
            'pixels_per_second': {'model': self.pixels_per_second, 'type': int},
            'scan_interval': {'model': self.scan_interval, 'type': float},
            'hold_duration': {'model': self.hold_duration, 'type': float},
            'loc_offset_variance': {'model': self.loc_offset_variance, 'type': int},
            'speed_variance': {'model': self.speed_variance, 'type': float},
            'hold_duration_variance': {'model': self.hold_duration_variance, 'type': float},
            'area_x1': {'model': self.area_x1, 'type': int},
            'area_y1': {'model': self.area_y1, 'type': int},
            'area_x2': {'model': self.area_x2, 'type': int},
            'area_y2': {'model': self.area_y2, 'type': int},
            'min_move_time': {'model': self.min_move_time, 'type': float},
            'max_move_time': {'model': self.max_move_time, 'type': float},
            'grid_visible': {'model': self.grid_visible, 'type': bool},
            'grid_latching': {'model': self.grid_latching, 'type': bool},
            'grid_spacing': {'model': self.grid_spacing, 'type': int},
            'grid_opacity': {'model': self.grid_opacity, 'type': float},
        }
        self.global_settings_ui_vars = {key: tk.StringVar() for key in self.global_settings_map}

        # --- Core Data Structures ---
        self.steps = []
        self.annotations = []
        self.template_cache = {}
        self.folder_image_cache = {}
        self.clipboard = []
        self.clipboard_origin = (0, 0)
        
        # --- UI State & Data ---
        self.selected_items = []
        self.properties_widgets = {}
        self._drag_data = {"start_x": 0, "start_y": 0, "item": None, "mode": "move", "initial_positions": []}
        self._marquee_data = {}
        self.area_overlays = {}
        self.zoom_factor = 1.0
        self.search_query = tk.StringVar()
        self.search_results = []
        self.current_search_index = -1
        self.search_query.trace_add('write', self._reset_search)

        # --- Execution State ---
        self.running = False
        self.current_step_index = 0
        self.current_step_start_time = 0
        self.executor_after_id = None
        self.delay_countdown_id = None
        self.timeout_countdown_id = None
        self.stop_requested = False
        self.start_step = tk.StringVar(value='1')
        self.automation_start_time = 0
        self.cycle_time_display = tk.StringVar(value="Cycle Time: 0.0s")
        self.cycle_time_updater_id = None
        
        # --- Hotkey / Capture Mode State ---
        self.f3_mode = None

        # --- Live Info & Logging ---
        self.last_detection_info = tk.StringVar(value="Detection: N/A")
        self.log_text = None 
        self.full_log_history = []
        self.log_search_query = tk.StringVar()
        self.log_search_query.trace_add('write', self.filter_log)
        self.log_auto_clear_lines = tk.IntVar(value=500)

        # --- Testing Panel ---
        self.active_test_type = tk.StringVar(value="PNG")
        self.test_png_mode = tk.StringVar(value='file')
        self.test_png_path = tk.StringVar()
        self.test_png_path_display = tk.StringVar(value="No path set")
        self.test_png_threshold = tk.DoubleVar(value=0.8)
        self.test_png_image_mode = tk.StringVar(value='Grayscale')
        self.test_png_count_expression = tk.StringVar(value='>= 1')
        self.test_area = None
        self.test_color_rgb = (255, 0, 0)
        self.test_color_tolerance = tk.IntVar(value=10)
        self.test_color_space = tk.StringVar(value='HSV')
        self.test_color_min_area = tk.IntVar(value=10)
        self.test_color_count_expression = tk.StringVar(value='>= 1')
        self.test_number_expression = tk.StringVar(value='> 0')
        self.psm_options = {
            "0: Orientation and script detection (OSD) only.": "0", "1: Automatic page segmentation with OSD.": "1",
            "3: Fully automatic page segmentation, but no OSD. (Default)": "3", "6: Assume a single uniform block of text.": "6",
            "7: Treat the image as a single text line.": "7", "8: Treat the image as a single word.": "8",
            "10: Treat the image as a single character.": "10", "13: Raw line. Treat the image as a single text line, bypassing hacks.": "13"
        }
        self.oem_options = { "0: Legacy Engine only.": "0", "1: Neural nets LSTM engine only.": "1", "2: Legacy + LSTM engines.": "2", "3: Default, based on what is available.": "3" }
        self.test_number_oem = tk.StringVar(value="3: Default, based on what is available.")
        self.test_number_psm = tk.StringVar(value="6: Assume a single uniform block of text.")

        # --- GE Interface ---
        self.api_headers = {'User-Agent': 'Flowchart Automation Tool - Contact on GitHub'}
        self.item_mapping_cache = None; self.item_price_cache = {}; self.all_item_prices_cache = None; self.hourly_volume_cache = None
        self.ge_interface_item_name = tk.StringVar(); self.ge_interface_item_quantity = tk.StringVar(value="1")
        self.ge_interface_buy_price_strategy = tk.StringVar(value='Flip-Buy (use Insta-Sell)')
        self.ge_interface_sell_price_strategy = tk.StringVar(value='Flip-Sell (use Insta-Buy)')
        self.ge_interface_buy_custom_price = tk.StringVar(value="0")
        self.ge_interface_sell_custom_price = tk.StringVar(value="0")
        self.ge_interface_buy_price_margin = tk.StringVar(value="1")
        self.ge_interface_sell_price_margin = tk.StringVar(value="1")
        self.ge_interface_display_buy_price = tk.StringVar(value="N/A")
        self.ge_interface_display_buy_total = tk.StringVar(value="N/A")
        self.ge_interface_display_sell_price = tk.StringVar(value="N/A")
        self.ge_interface_display_sell_total = tk.StringVar(value="N/A")
        self.ge_interface_last_data = None
        self.ge_auto_update_enabled = tk.BooleanVar(value=False)
        self.ge_auto_update_enabled.trace_add('write', self._toggle_ge_auto_update)
        self.ge_auto_update_interval = tk.StringVar(value="60"); self.ge_auto_update_after_id = None

        # --- Threading Lock for Detection ---
        self.detection_lock = threading.Lock()
        self.detection_thread = None
        self.detection_result = None

        # --- Final UI Setup ---
        self.build_ui()
        self.setup_hotkeys()
        self.apply_theme()
        self.log("Application initialized successfully.")
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)


if __name__ == "__main__":
    root = tk.Tk()
    app = FlowchartClickerApp(root)
    root.mainloop()
