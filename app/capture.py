import tkinter as tk
from tkinter import filedialog, messagebox
import pyautogui
import time
import os
import keyboard

class CaptureMixin:
    def enter_f3_mode(self, action):
        if self.running: return
        index = self.selected_items[0]['index'] if (self.selected_items and len(self.selected_items) == 1 and self.selected_items[0]['type'] == 'step') else None
        context = 'test' if action in ['pick_test_color'] else 'step'
        if context == 'step' and index is None: return
        if self.hide_on_select.get(): self.root.withdraw()
        self.f3_mode = {'action': action, 'index': index, 'context': context}
        
        action_text_map = {'pick_color': "PICKING COLOR", 'pick_location': "GETTING LOCATION"}
        action_text = action_text_map.get(action, "CAPTURING")
        
        self.status_label_color_state = 'orange'; self.status_label.config(text=f"{action_text}: Move mouse and press F3", foreground=self.current_theme['status_orange']); self.log(f"Entering picker mode. Press F3 to capture.", "orange")

    def capture_from_hotkey(self):
        if self.f3_mode is None: return
        try:
            context, index, action = self.f3_mode['context'], self.f3_mode['index'], self.f3_mode['action']; x, y = pyautogui.position()
            if context == 'step':
                step = self.steps[index]
                if action == 'pick_color': 
                    step['rgb'] = pyautogui.pixel(x, y)
                    if step.get('pixel_detect_enabled'):
                        step['pixel_coords'] = (x, y)
                    self.log(f"Captured color {step['rgb']} for Step {index + 1}.")
                elif action == 'pick_location': 
                    step['coords'] = (x, y); self.log(f"Captured location {(x,y)} for Step {index + 1}.")
                self.populate_properties_panel()
            elif context == 'test':
                if action == 'pick_test_color': 
                    self.test_color_rgb = pyautogui.pixel(x, y)
                    hex_color = self.rgb_to_hex(self.test_color_rgb)
                    self.test_color_swatch.config(bg=hex_color)
                    if hasattr(self, 'test_color_swatch_count'):
                        self.test_color_swatch_count.config(bg=hex_color)
                    self.log(f"Captured color {self.test_color_rgb} for testing.")
        except Exception as e: messagebox.showerror("Error", f"Could not capture: {e}")
        finally:
            if self.hide_on_select.get():
                self.root.config(bg=self.current_theme['bg'])
                self.root.deiconify()
            self.f3_mode = None; self.status_label_color_state = 'blue'; self.status_label.config(text="Status: Stopped", foreground=self.current_theme.get('status_blue', 'blue'))

    def select_area_for_step(self): 
        if self.selected_items and len(self.selected_items) == 1 and self.selected_items[0]['type'] == 'step': 
            self.select_area_mode(step_index=self.selected_items[0]['index'])
    
    def select_area_for_test(self): self.select_area_mode(is_test=True)

    def set_global_area_to_fullscreen(self):
        screen_w, screen_h = pyautogui.size()
        self.global_settings_ui_vars['area_x1'].set(0); self.global_settings_ui_vars['area_y1'].set(0)
        self.global_settings_ui_vars['area_x2'].set(screen_w); self.global_settings_ui_vars['area_y2'].set(screen_h)
        self.apply_global_settings(); self.log("Set global area to full screen.")

    def set_step_area_to_fullscreen(self):
        if not (self.selected_items and len(self.selected_items) == 1 and self.selected_items[0]['type'] == 'step'): return
        screen_w, screen_h = pyautogui.size(); index = self.selected_items[0]['index']; self.steps[index]['area'] = (0, 0, screen_w, screen_h); self.log(f"Set area for Step {index + 1} to full screen."); self.populate_properties_panel()
        self.update_all_area_overlays()

    def set_step_area_to_global(self):
        """Sets the selected step's area to None, causing it to use the global area."""
        if not (self.selected_items and len(self.selected_items) == 1 and self.selected_items[0]['type'] == 'step'):
            return
        index = self.selected_items[0]['index']
        self.steps[index]['area'] = None
        self.log(f"Set Step {index + 1} to use the Global Area setting.")
        self.populate_properties_panel()  # Refresh UI to show "Area: Global"
        self.update_all_area_overlays() # Update visual overlays

    def browse_for_step(self):
        if not (self.selected_items and len(self.selected_items) == 1 and self.selected_items[0]['type'] == 'step'): return
        index = self.selected_items[0]['index']; step = self.steps[index]
        mode = self.properties_widgets['png_mode'].get()
        path = filedialog.askopenfilename(filetypes=[("PNG Files", "*.png")]) if mode == 'file' else filedialog.askdirectory()
        
        if path:
            old_path = step.get('path', '')
            if old_path != path:
                image_mode = step.get('image_mode', 'Grayscale')
                old_single_key = f"{old_path}|{image_mode}"
                if old_single_key in self.template_cache:
                    del self.template_cache[old_single_key]
                old_folder_key = f"{old_path}|{image_mode}"
                if old_folder_key in self.folder_image_cache:
                    del self.folder_image_cache[old_folder_key]

            step['path'] = path
            self.properties_widgets['path'].config(text=os.path.basename(path), fg=self.current_theme['fg'])
            self._update_png_preview(step)
            self.log(f"Set path for Step {index + 1} to '{os.path.basename(path)}'.")

    def snip_image_for_step(self, event=None):
        if self.running or not (self.selected_items and len(self.selected_items) == 1 and self.selected_items[0]['type'] == 'step'): return
        index = self.selected_items[0]['index']
        window_state, current_geometry = self.root.state(), self.root.geometry()
        
        if self.hide_on_select.get():
            self.root.withdraw()
            self.root.update_idletasks()
            time.sleep(0.5) # Wait for OS to redraw behind the withdrawn window
            
        overlay = tk.Toplevel(self.root); overlay.attributes("-alpha", 0.3, "-topmost", True); overlay.overrideredirect(True); w, h = self.root.winfo_screenwidth(), self.root.winfo_screenheight(); overlay.geometry(f"{w}x{h}+0+0"); canvas = tk.Canvas(overlay, cursor="cross", bg="grey10"); canvas.pack(fill=tk.BOTH, expand=True); rect_id, start_x, start_y = None, 0, 0
        
        def on_press(e): nonlocal start_x, start_y, rect_id; start_x, start_y = e.x_root, e.y_root; rect_id = canvas.create_rectangle(e.x, e.y, e.x, e.y, outline="#3399ff", width=2, dash=(4,2))
        def on_drag(e):
            if rect_id: canvas.coords(rect_id, start_x, start_y, e.x_root, e.y_root)
        
        def on_release(e):
            x1, y1, x2, y2 = min(start_x, e.x_root), min(start_y, e.y_root), max(start_x, e.x_root), max(start_y, e.y_root)
            overlay.destroy()

            # Abort if area is too small, but make sure to restore window first
            if (x2 - x1) < 10 or (y2 - y1) < 10:
                if self.hide_on_select.get():
                    self.root.deiconify()
                    self.root.state('zoomed') if window_state == 'zoomed' else self.root.geometry(current_geometry)
                self.root.focus_force()
                self.log("Snipping cancelled: area was too small.", "orange")
                return

            # Capture the screen region while the main window is still hidden
            captured_image = None
            try:
                time.sleep(0.1) # Brief pause for overlay to vanish
                captured_image = pyautogui.screenshot(region=(x1, y1, x2 - x1, y2 - y1))
            except Exception as ex:
                self.log(f"Error capturing screen snippet: {ex}", "red")
            
            # Now that the capture is done, restore the main window
            if self.hide_on_select.get():
                self.root.deiconify()
                self.root.state('zoomed') if window_state == 'zoomed' else self.root.geometry(current_geometry)
            self.root.focus_force()

            # If capture was successful, ask the user where to save it
            if captured_image:
                filepath = filedialog.asksaveasfilename(title="Save Snippet As", defaultextension=".png", filetypes=[("PNG Files", "*.png")], initialfile=f"snippet_step_{index+1}.png")
                if filepath:
                    try:
                        captured_image.save(filepath)
                        self.steps[index]['path'] = filepath
                        self.populate_properties_panel()
                        self._update_png_preview(self.steps[index])
                        self.log(f"Saved snippet and set path for Step {index+1}.")
                    except Exception as ex:
                        messagebox.showerror("Save Error", f"Failed to save snippet: {ex}")
                        self.log(f"Error saving snippet: {ex}", "red")

        canvas.bind("<ButtonPress-1>", on_press)
        canvas.bind("<B1-Motion>", on_drag)
        canvas.bind("<ButtonRelease-1>", on_release)

    def browse_for_test_path(self):
        mode = self.test_png_mode.get(); path = filedialog.askopenfilename(filetypes=[("PNG Files", "*.png")]) if mode == 'file' else filedialog.askdirectory()
        if path: self.test_png_path.set(path); self.test_png_path_display.set(os.path.basename(path))

    def snip_image_for_test(self):
        if self.running: return
        window_state, current_geometry = self.root.state(), self.root.geometry()

        if self.hide_on_select.get():
            self.root.withdraw()
            self.root.update_idletasks()
            time.sleep(0.5) # Wait for OS to redraw behind the withdrawn window

        overlay = tk.Toplevel(self.root); overlay.attributes("-alpha", 0.3, "-topmost", True); overlay.overrideredirect(True); w, h = self.root.winfo_screenwidth(), self.root.winfo_screenheight(); overlay.geometry(f"{w}x{h}+0+0"); canvas = tk.Canvas(overlay, cursor="cross", bg="grey10"); canvas.pack(fill=tk.BOTH, expand=True); rect_id, start_x, start_y = None, 0, 0
        
        def on_press(e): nonlocal start_x, start_y, rect_id; start_x, start_y = e.x_root, e.y_root; rect_id = canvas.create_rectangle(e.x, e.y, e.x, e.y, outline="#3399ff", width=2, dash=(4,2))
        def on_drag(e):
            if rect_id: canvas.coords(rect_id, start_x, start_y, e.x_root, e.y_root)

        def on_release(e):
            x1, y1, x2, y2 = min(start_x, e.x_root), min(start_y, e.y_root), max(start_x, e.x_root), max(start_y, e.y_root)
            overlay.destroy()

            # Abort if area is too small, but make sure to restore window first
            if (x2 - x1) < 10 or (y2 - y1) < 10:
                if self.hide_on_select.get():
                    self.root.deiconify()
                    self.root.state('zoomed') if window_state == 'zoomed' else self.root.geometry(current_geometry)
                self.root.focus_force()
                self.log("Snipping cancelled: area was too small.", "orange")
                return

            # Capture the screen region while the main window is still hidden
            captured_image = None
            try:
                time.sleep(0.1) # Brief pause for overlay to vanish
                captured_image = pyautogui.screenshot(region=(x1, y1, x2 - x1, y2 - y1))
            except Exception as ex:
                self.log(f"Error capturing screen snippet: {ex}", "red")

            # Now that the capture is done, restore the main window
            if self.hide_on_select.get():
                self.root.deiconify()
                self.root.state('zoomed') if window_state == 'zoomed' else self.root.geometry(current_geometry)
            self.root.focus_force()
            
            # If capture was successful, ask the user where to save it
            if captured_image:
                filepath = filedialog.asksaveasfilename(title="Save Test Snippet As", defaultextension=".png", filetypes=[("PNG Files", "*.png")], initialfile="test_snippet.png")
                if filepath:
                    try:
                        captured_image.save(filepath)
                        self.test_png_path.set(filepath)
                        self.test_png_path_display.set(os.path.basename(filepath))
                        self.log(f"Saved snippet and set path for PNG test.")
                    except Exception as ex:
                        messagebox.showerror("Save Error", f"Failed to save snippet: {ex}")
                        self.log(f"Error saving snippet: {ex}", "red")
        
        canvas.bind("<ButtonPress-1>", on_press)
        canvas.bind("<B1-Motion>", on_drag)
        canvas.bind("<ButtonRelease-1>", on_release)

    def select_area_mode(self, step_index=None, is_test=False):
        if self.running: return
        window_state, current_geometry = self.root.state(), self.root.geometry()
        if self.hide_on_select.get():
            self.root.withdraw()
            self.root.update_idletasks()
            time.sleep(0.3)
        overlay = tk.Toplevel(self.root); overlay.attributes("-alpha", 0.25, "-topmost", True); overlay.overrideredirect(True); w, h = self.root.winfo_screenwidth(), self.root.winfo_screenheight(); overlay.geometry(f"{w}x{h}+0+0"); canvas = tk.Canvas(overlay, cursor="cross", bg="grey10"); canvas.pack(fill=tk.BOTH, expand=True); rect_id, start_x, start_y = None, 0, 0
        def on_press(e): nonlocal start_x, start_y, rect_id; start_x, start_y = e.x_root, e.y_root; rect_id = canvas.create_rectangle(e.x, e.y, e.x, e.y, outline="red", width=2)
        def on_drag(e):
            if rect_id: canvas.coords(rect_id, start_x, start_y, e.x_root, e.y_root)
        def on_release(e):
            x1, y1, x2, y2 = min(start_x, e.x_root), min(start_y, e.y_root), max(start_x, e.x_root), max(start_y, e.y_root); overlay.destroy()
            if self.hide_on_select.get():
                self.root.config(bg=self.current_theme['bg'])
                self.root.deiconify()
                self.root.state('zoomed') if window_state == 'zoomed' else self.root.geometry(current_geometry)
            if (x2-x1)>10 and (y2-y1)>10:
                area = (x1, y1, x2, y2)
                if is_test: 
                    self.test_area = area
                    area_text = f"Area: {x2-x1}x{y2-y1}"
                    for btn in self.test_area_buttons:
                        if btn.winfo_exists():
                            btn.config(text=area_text)
                    self.log("Set area for testing.")
                elif step_index is not None: 
                    self.steps[step_index]['area'] = area; self.populate_properties_panel(); self.log(f"Set area for Step {step_index+1}.")
                    self.update_all_area_overlays()
                else: 
                    self.global_settings_ui_vars['area_x1'].set(x1); self.global_settings_ui_vars['area_y1'].set(y1); self.global_settings_ui_vars['area_x2'].set(x2); self.global_settings_ui_vars['area_y2'].set(y2)
                    self.apply_global_settings(); self.log("Set global area.")
        canvas.bind("<ButtonPress-1>", on_press); canvas.bind("<B1-Motion>", on_drag); canvas.bind("<ButtonRelease-1>", on_release)


