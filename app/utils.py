import tkinter as tk
from tkinter import scrolledtext
import cv2
import numpy as np
import pyautogui
import time
import os
import math
import keyboard
import threading
try:
    import pytesseract
    from PIL import Image
except ImportError:
    pytesseract = None

class UtilsMixin:
    def on_closing(self):
        self.destroy_all_overlays()
        self._stop_ge_auto_updater()
        if self.running: self.stop()
        self.root.destroy()

    def log(self, message, color_name=None):
        theme = self.current_theme
        if color_name in ["green", "orange", "red"]: self.status_label_color_state = color_name; self.status_label.config(foreground=theme[f'status_{color_name}'])
        log_entry = f"[{time.strftime('%H:%M:%S')}] {message}"; self.full_log_history.append(log_entry); self.filter_log()

    def log_execution(self, message, color_name=None):
        """Logs a message during script execution, respecting the step's log setting."""
        if not self.running or not (0 <= self.current_step_index < len(self.steps)):
            self.log(message, color_name)
            return

        step = self.steps[self.current_step_index]
        if step.get('enable_logging', True):
            self.log(message, color_name)

    def filter_log(self, *args):
        query = self.log_search_query.get().lower(); self.log_text.config(state='normal'); self.log_text.delete('1.0', tk.END)
        filtered_log = [line for line in self.full_log_history if query in line.lower()] if query else self.full_log_history
        for line in filtered_log: self.log_text.insert(tk.END, line + "\n")
        max_lines = self.log_auto_clear_lines.get()
        if max_lines > 0 and not query:
            num_lines = int(self.log_text.index('end-1c').split('.')[0])
            if num_lines > max_lines:
                self.log_text.delete('1.0', f'{num_lines - max_lines + 1}.0')
                self.full_log_history = self.full_log_history[-(max_lines):]
        self.log_text.config(state='disabled'); self.log_text.yview(tk.END)

    def clear_log(self):
        self.full_log_history.clear(); self.log_search_query.set(""); self.log("Log cleared.")

    def get_node_center(self, index):
        step, z = self.steps[index], self.zoom_factor
        return (step.get('x', 50) + step.get('_width', 180*z)/z/2, step.get('y', 50) + step.get('_height', 60*z)/z/2)

    def run_test(self):
        def log_test(message): self.test_results_text.config(state='normal'); self.test_results_text.insert(tk.END, message + "\n"); self.test_results_text.config(state='disabled'); self.test_results_text.yview(tk.END)
        self.test_results_text.config(state='normal'); self.test_results_text.delete('1.0', tk.END); self.test_results_text.config(state='disabled')
        
        active_type = self.active_test_type.get()
        area = self.test_area or (self.area_x1.get(), self.area_y1.get(), self.area_x2.get(), self.area_y2.get()); w, h = area[2]-area[0], area[3]-area[1]
        if w < 1 or h < 1: log_test("ERROR: Test area is invalid."); return
        log_test(f"--- Running {self.active_test_type.get()} Test ---\nUsing area: {area}")
        try:
            if self.hide_on_select.get(): window_state, current_geometry = self.root.state(), self.root.geometry(); self.root.withdraw(); time.sleep(0.3)
            screen_cv = cv2.cvtColor(np.array(pyautogui.screenshot(region=(area[0], area[1], w, h))), cv2.COLOR_RGB2BGR)
        except Exception as e: log_test(f"ERROR: Failed to capture screen: {e}"); return
        finally:
            if self.hide_on_select.get(): 
                self.root.config(bg=self.current_theme['bg'])
                self.root.deiconify(); 
                self.root.state('zoomed') if window_state == 'zoomed' else self.root.geometry(current_geometry)
        
        if active_type == 'PNG':
            path, threshold, mode, image_mode = self.test_png_path.get(), self.test_png_threshold.get(), self.test_png_mode.get(), self.test_png_image_mode.get()
            log_test(f"Searching with Image Mode: {image_mode}")
            if not path: log_test("ERROR: No PNG path selected."); return
            image_paths = []
            if mode == 'file' and os.path.isfile(path): image_paths.append(path)
            elif mode == 'folder' and os.path.isdir(path): image_paths = [os.path.join(path, f) for f in os.listdir(path) if f.lower().endswith('.png')]
            if not image_paths: log_test("No valid PNG images found."); return
            if image_mode == 'Grayscale': screen_processed = cv2.cvtColor(screen_cv, cv2.COLOR_BGR2GRAY)
            elif image_mode == 'Binary (B&W)': _, screen_processed = cv2.threshold(cv2.cvtColor(screen_cv, cv2.COLOR_BGR2GRAY), 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
            else: screen_processed = screen_cv
            found_count = 0
            for fpath in image_paths:
                template_data = self.load_template(fpath, image_mode)
                if template_data[0] is None: log_test(f"Warning: Could not load template {os.path.basename(fpath)}"); continue
                match = self.find_template_in_region(screen_processed, area[0:2], template_data, threshold)
                if match and math.isfinite(match[2]):
                    found_count += 1; log_test(f"SUCCESS: Found {os.path.basename(fpath)} with {match[2]*100:.2f}% confidence.")
            if found_count == 0: log_test("RESULT: No PNGs detected above the threshold.")
            else: log_test(f"--- Finished: Found {found_count} matching PNG(s). ---")
        elif active_type == 'Color':
            rgb, tolerance, space = self.test_color_rgb, self.test_color_tolerance.get(), self.test_color_space.get()
            log_test(f"Searching for RGB: {rgb} with tolerance: {tolerance} in {space} space")
            mock_step = {'rgb': rgb, 'tolerance': tolerance, 'min_pixel_area': 10}
            if space == 'RGB': 
                target_pos, contour_area = self.find_color_on_screen_rgb(screen_cv, area[0:2], mock_step)
            else: 
                target_pos, contour_area = self.find_color_on_screen_hsv(screen_cv, area[0:2], mock_step)
            if target_pos: log_test(f"SUCCESS: Color found at {target_pos}.\nDetected area size: {contour_area:.2f} pixels.")
            else: log_test("RESULT: Color not detected.")
        elif active_type == 'PNG Count':
            path, threshold, mode, image_mode, expression_str = self.test_png_path.get(), self.test_png_threshold.get(), self.test_png_mode.get(), self.test_png_image_mode.get(), self.test_png_count_expression.get()
            log_test(f"Counting with Image Mode: {image_mode}")
            if not path: log_test("ERROR: No PNG path selected."); return
            
            mock_step = { 'path': path, 'threshold': threshold, 'mode': mode, 'image_mode': image_mode }
            count = self.find_and_count_png(screen_cv, area[0:2], mock_step)
            log_test(f"Found {count} instances.")
            
            try:
                expression = expression_str.split()
                if len(expression) == 2:
                    op, val = expression[0], int(expression[1])
                    op_map = {'>': count > val, '<': count < val, '>=': count >= val, '<=': count <= val, '==': count == val, '!=': count != val}
                    result = op_map.get(op, False)
                    
                    if result: log_test(f"SUCCESS: Condition '{count} {op} {val}' is TRUE.")
                    else: log_test(f"RESULT: Condition '{count} {op} {val}' is FALSE.")
                else: raise ValueError("Expression must have 2 parts (e.g., '>= 5')")
            except (ValueError, IndexError) as e: log_test(f"ERROR: Invalid expression format.\nDetails: {e}")
        elif active_type == 'Color Count':
            rgb, tolerance, space = self.test_color_rgb, self.test_color_tolerance.get(), self.test_color_space.get()
            min_area = self.test_color_min_area.get()
            expression_str = self.test_color_count_expression.get()
            log_test(f"Counting Color RGB: {rgb} with tolerance: {tolerance} in {space} space")
            log_test(f"Min Area: {min_area}px, Expression: '{expression_str}'")
            mock_step = {'rgb': rgb, 'tolerance': tolerance, 'min_pixel_area': min_area, 'color_space': space}
            count = self.find_and_count_color(screen_cv, area[0:2], mock_step)
            log_test(f"Found {count} color blobs.")
            try:
                expression = expression_str.split()
                if len(expression) != 2: raise ValueError("Expression must have 2 parts (e.g., '>= 5')")
                op, val = expression[0], int(expression[1])
                op_map = {'>': count > val, '<': count < val, '>=': count >= val, '<=': count <= val, '==': count == val, '!=': count != val}
                result = op_map.get(op, False)
                if result: log_test(f"SUCCESS: Condition '{count} {op} {val}' is TRUE.")
                else: log_test(f"RESULT: Condition '{count} {op} {val}' is FALSE.")
            except (ValueError, IndexError) as e: log_test(f"ERROR: Invalid expression format.\nDetails: {e}")
        elif active_type == 'Number':
            expression = self.test_number_expression.get()
            psm_mode = self.psm_options.get(self.test_number_psm.get(), '6')
            oem_mode = self.oem_options.get(self.test_number_oem.get(), '3')
            log_test(f"Searching for number: '{expression}' (OEM: {oem_mode}, PSM: {psm_mode})")
            preprocessed_img = self.preprocess_for_ocr(screen_cv)
            ocr_config = f'--oem {oem_mode} --psm {psm_mode} -c tessedit_char_whitelist=0123456789:;,.-'
            try:
                ocr_text = pytesseract.image_to_string(Image.fromarray(preprocessed_img), config=ocr_config)
                cleaned_text = "".join(filter(lambda x: x in '0123456789.-', ocr_text)); log_test(f"OCR detected: '{cleaned_text}'")
                try:
                    num = float(cleaned_text); expr = expression.split(); op, val = expr[0], float(expr[1])
                    op_map = {'>': num > val, '<': num < val, '>=': num >= val, '<=': num <= val, '==': num == val, '!=': num != val}
                    if op_map.get(op, False): log_test(f"SUCCESS: {num} {op} {val} is TRUE.")
                    else: log_test(f"RESULT: {num} {op} {val} is FALSE.")
                except (ValueError, IndexError) as e: log_test(f"ERROR: Invalid expression format.\nDetails: {e}")
            except pytesseract.TesseractNotFoundError: log_test("ERROR: Tesseract OCR not found.")
            except Exception as e: log_test(f"ERROR during OCR: {e}")


