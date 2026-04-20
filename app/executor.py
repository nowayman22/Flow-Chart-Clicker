import tkinter as tk
from tkinter import messagebox
import time
import threading
import os
import cv2
import numpy as np
import pyautogui

class ExecutorMixin:
    def _pre_cache_folder_templates(self):
        """
        Proactively loads and caches templates for all PNG steps in folder mode.
        This prevents a long delay on the first detection attempt of each step.
        """
        self.log("Pre-caching templates for PNG Folder steps...")
        count = 0
        for i, step in enumerate(self.steps):
            if step.get('type') == 'png' and step.get('mode') == 'folder' and step.get('path') and os.path.isdir(step['path']):
                image_mode = step.get('image_mode', 'Grayscale')
                folder_cache_key = f"{step['path']}|{image_mode}"
                
                if folder_cache_key not in self.folder_image_cache:
                    self.folder_image_cache[folder_cache_key] = []
                    image_paths = [os.path.join(step['path'], fname) for fname in os.listdir(step['path']) if fname.lower().endswith('.png')]
                    for fpath in image_paths:
                        template_data = self.load_template(fpath, image_mode)
                        if template_data[0] is not None:
                            self.folder_image_cache[folder_cache_key].append(template_data)
                    
                    num_cached = len(self.folder_image_cache[folder_cache_key])
                    if num_cached > 0:
                        self.log(f" > Cached {num_cached} templates for Step {i+1} from '{os.path.basename(step['path'])}'.")
                        count += num_cached
        if count > 0:
            self.log(f"Finished pre-caching {count} total templates.", "green")
        else:
            self.log("No new PNG Folder steps found to pre-cache.")

    def start(self):
        if self.running or self.f3_mode: return
        if not self.steps: messagebox.showerror("Error", "No steps defined."); return
        
        self._pre_cache_folder_templates()
        
        resetted_items = []
        for i, step in enumerate(self.steps):
            if step.get('type') == 'logical' and step.get('reset_on_start'):
                if step.get('logical_type') == 'Count':
                    step['counter_value'] = 0
                    resetted_items.append(f"Counter for Step {i+1}")
                elif step.get('logical_type') == 'Wait':
                    step['timer_start_time'] = None
                    step['last_cycle_time'] = 'N/A'
                    resetted_items.append(f"Timer for Step {i+1}")
                elif step.get('logical_type') == 'Movement Detect':
                    step['_previous_frame_for_movement'] = None
                    resetted_items.append(f"Movement Comparison for Step {i+1}")

        if resetted_items: self.log(f"Reset on start: {', '.join(resetted_items)}.")
        try:
            start_index = int(self.start_step.get()) - 1
            if not (0 <= start_index < len(self.steps)): messagebox.showerror("Invalid Start Step", f"Start step must be between 1 and {len(self.steps)}."); return
            self.current_step_index = start_index
        except ValueError: messagebox.showerror("Invalid Input", "Start step must be a valid number."); return
        
        # --- FIX: Set running flag to True BEFORE starting the timer loop ---
        self.running = True
        self.automation_start_time = time.time()
        self.cycle_time_display.set("Cycle Time: 00:00:00")
        self._update_cycle_time()
        
        self.status_label_color_state = 'green'
        self.log("Automation started.", 'green')
        self.start_btn.config(state=tk.DISABLED)
        self.stop_btn.config(state=tk.NORMAL)
        self.advance_step()

    def stop(self, message="Status: Stopped", color_state='blue'):
        """
        Stops the automation script. This method is designed to be thread-safe,
        as it can be called from the keyboard listener thread.
        """
        # Use a flag to prevent re-entry from multiple rapid presses
        if self.stop_requested:
            return
        self.stop_requested = True
        
        # 1. Immediately set the main running flag to False. This is the primary
        #    mechanism to halt the execution loops and interruptible moves.
        self.running = False
        
        # 2. Cancel any pending `after` calls, which schedule future work.
        #    This is thread-safe according to Tkinter's documentation.
        if self.executor_after_id: self.root.after_cancel(self.executor_after_id)
        if self.delay_countdown_id: self.root.after_cancel(self.delay_countdown_id)
        if self.timeout_countdown_id: self.root.after_cancel(self.timeout_countdown_id)
        if self.cycle_time_updater_id: self.root.after_cancel(self.cycle_time_updater_id); self.cycle_time_updater_id = None

        # 3. Safely handle the background detection thread state.
        with self.detection_lock:
            self.detection_thread = None
            self.detection_result = None

        # 4. Schedule the final state changes and UI updates to run in the main Tkinter thread.
        self.root.after(0, self._finalize_stop_ui, message, color_state)

    def _finalize_stop_ui(self, message, color_state):
        """
        Performs the final, non-thread-safe actions to complete the stop process.
        This method MUST be called from the main Tkinter thread via `root.after()`.
        """
        if self.automation_start_time > 0:
            final_elapsed = time.time() - self.automation_start_time
            total_seconds = int(final_elapsed)
            hours = total_seconds // 3600
            minutes = (total_seconds % 3600) // 60
            seconds = total_seconds % 60
            time_str = f"{hours:02}:{minutes:02}:{seconds:02}"
            self.log(f"Automation stopped. Total cycle time: {time_str}.")
            self.automation_start_time = 0

        # Update the 'Start Step' field if the option is enabled.
        if self.start_at_stopped_pos.get() and self.current_step_index < len(self.steps):
            next_start_step = str(self.current_step_index + 1)
            self.start_step.set(next_start_step)
            self.log(f"Next start step set to {next_start_step}.")

        # Reset internal state variables that hold `after` IDs.
        self.executor_after_id = None
        self.delay_countdown_id = None
        self.timeout_countdown_id = None

        # Update all UI elements to reflect the stopped state.
        self.start_btn.config(state=tk.NORMAL)
        self.stop_btn.config(state=tk.DISABLED)
        self.status_label_color_state = color_state
        color = self.current_theme.get(f"status_{color_state}", self.current_theme.get('status_blue'))
        self.status_label.config(text=message, foreground=color)
        self.delay_countdown_label.config(text="")
        self.timeout_countdown_label.config(text="")
        self.redraw_flowchart() # Redraw to remove the 'current step' highlight
        
        # Reset the request flag after everything is done.
        self.stop_requested = False

    def advance_step(self):
        if not self.running or not self.steps: self.stop(); return
        if self.timeout_countdown_id: self.root.after_cancel(self.timeout_countdown_id); self.timeout_countdown_id = None
        self.timeout_countdown_label.config(text=""); self.last_detection_info.set("Detection: N/A")
        
        with self.detection_lock:
            self.detection_thread = None
            self.detection_result = None

        if self.current_step_index >= len(self.steps): self.log("Completed all steps.", "green"); self.stop("Status: Completed all steps", color_state='green'); return
        self.current_step_start_time = time.time(); self.redraw_flowchart(); self.run_step_executor()

    def _update_cycle_time(self):
        if self.running:
            elapsed = time.time() - self.automation_start_time
            total_seconds = int(elapsed)
            hours = total_seconds // 3600
            minutes = (total_seconds % 3600) // 60
            seconds = total_seconds % 60
            time_str = f"{hours:02}:{minutes:02}:{seconds:02}"
            self.cycle_time_display.set(f"Cycle Time: {time_str}")
            self.cycle_time_updater_id = self.root.after(100, self._update_cycle_time)

    def _perform_png_detection_in_thread(self, screen_cv, offset, step):
        """
        Runs the find_png method in a separate thread to avoid blocking the GUI.
        The result is stored in self.detection_result.
        """
        try:
            target_pos, confidence = self.find_png(screen_cv, offset, step)
            with self.detection_lock:
                if self.detection_thread is threading.current_thread():
                    self.detection_result = ('png', target_pos, confidence)
        except Exception as e:
            print(f"Error in detection thread: {e}")
            with self.detection_lock:
                if self.detection_thread is threading.current_thread():
                    self.detection_result = ('png', None, 0)

    def _perform_color_detection_in_thread(self, screen_cv, offset, step):
        """
        Runs color detection in a separate thread. The result is stored in self.detection_result.
        """
        try:
            target_pos, contour_area = None, 0
            if step.get('pixel_detect_enabled', False):
                coords = step.get('pixel_coords')
                if coords:
                    current_rgb = pyautogui.pixel(coords[0], coords[1])
                    target_rgb = step.get('rgb'); tolerance = step.get('tolerance'); color_space = step.get('color_space', 'HSV')
                    match = False
                    if color_space == 'RGB':
                        match = all(abs(c1 - c2) <= tolerance for c1, c2 in zip(current_rgb, target_rgb))
                    else: 
                        current_bgr = np.uint8([[list(reversed(current_rgb))]])
                        target_bgr = np.uint8([[list(reversed(target_rgb))]])
                        current_hsv = cv2.cvtColor(current_bgr, cv2.COLOR_BGR2HSV)[0][0]
                        target_hsv = cv2.cvtColor(target_bgr, cv2.COLOR_BGR2HSV)[0][0]
                        h, s, v = int(target_hsv[0]), int(target_hsv[1]), int(target_hsv[2])
                        h_tol, s_tol, v_tol = int(tolerance*1.8), int(tolerance*2.5), int(tolerance*2.5)
                        h_diff = abs(int(current_hsv[0]) - h); h_match = min(h_diff, 180 - h_diff) <= h_tol
                        s_match = abs(int(current_hsv[1]) - s) <= s_tol
                        v_match = abs(int(current_hsv[2]) - v) <= v_tol
                        match = h_match and s_match and v_match
                    if match:
                        target_pos, contour_area = coords, 1 
            else: 
                color_space = step.get('color_space', 'HSV')
                if color_space == 'RGB':
                    target_pos, contour_area = self.find_color_on_screen_rgb(screen_cv, offset, step)
                else:
                    target_pos, contour_area = self.find_color_on_screen_hsv(screen_cv, offset, step)
            
            with self.detection_lock:
                if self.detection_thread is threading.current_thread():
                    self.detection_result = ('color', target_pos, contour_area)
        except Exception as e:
            print(f"Error in color detection thread: {e}")
            with self.detection_lock:
                if self.detection_thread is threading.current_thread():
                    self.detection_result = ('color', None, 0)

    def _perform_movement_detection_in_thread(self, current_frame_cv, previous_frame, step):
        """
        Runs movement comparison logic in a separate thread.
        """
        try:
            if previous_frame.shape != current_frame_cv.shape:
                raise ValueError("Frame dimension mismatch during movement detection.")

            diff = cv2.absdiff(previous_frame, current_frame_cv)
            _, thresholded_diff = cv2.threshold(diff, 30, 255, cv2.THRESH_BINARY)
            
            non_zero_count = np.count_nonzero(thresholded_diff)
            total_pixels = thresholded_diff.size
            change_percentage = (non_zero_count / total_pixels) * 100 if total_pixels > 0 else 0
            
            tolerance = step.get('movement_tolerance', 5.0)
            is_still = change_percentage <= tolerance

            with self.detection_lock:
                if self.detection_thread is threading.current_thread():
                    self.detection_result = ('movement', is_still, change_percentage)
        except Exception as e:
            print(f"Error in movement detection thread: {e}")
            with self.detection_lock:
                if self.detection_thread is threading.current_thread():
                     self.detection_result = ('movement', False, -1) # Indicate error

    def run_step_executor(self):
        if not self.running: return
        if not (0 <= self.current_step_index < len(self.steps)):
            self.log(f"Error: Invalid step index {self.current_step_index} detected. Stopping.", "red"); self.stop("Status: Stopped due to invalid index", "red"); return

        step = self.steps[self.current_step_index]
        step['_last_run_info'] = {'timestamp': time.time(), 'result': 'Running', 'details': 'Executing...'}
        self.status_label.config(text=f"Running Step {self.current_step_index + 1}: {step.get('name', '')}", foreground=self.current_theme['status_green'])
        
        is_timeout_step = (step.get('type') in ['color', 'png']) or \
                          (step.get('logical_type') in ['Number', 'Movement Detect'])

        if is_timeout_step and step.get('timeout', 0) > 0:
            if time.time() - self.current_step_start_time > step['timeout']: self.handle_timeout(); return
            self.update_timeout_countdown(self.current_step_start_time, step['timeout'])
            
        target_pos, step_succeeded = None, False
        try:
            if step['type'] == 'color' and step.get('action') == 'Color Count':
                area = step.get('area') or (self.area_x1.get(), self.area_y1.get(), self.area_x2.get(), self.area_y2.get())
                w, h = area[2] - area[0], area[3] - area[1]
                if w < 1 or h < 1:
                    self.log_execution(f"Step {self.current_step_index + 1}: Invalid area for Color Count. Failing.", "red")
                    self.handle_timeout(); return

                screenshot = pyautogui.screenshot(region=(area[0], area[1], w, h))
                screen_cv = cv2.cvtColor(np.array(screenshot), cv2.COLOR_RGB2BGR)
                
                count = self.find_and_count_color(screen_cv, area[0:2], step)
                expression_str = step.get('count_expression', '>= 1')
                
                step.setdefault('_count_current_cycle', 0)
                max_cycles = step.get('count_max_cycles', 1)
                
                self.last_detection_info.set(f"Color Count: Found {count} blobs. Condition: {expression_str}")
                
                try:
                    expression = expression_str.split()
                    if len(expression) != 2: raise ValueError("Expression must have 2 parts (e.g., '>= 5')")
                    op, val = expression[0], int(expression[1])
                    op_map = {'>': count > val, '<': count < val, '>=': count >= val, '<=': count <= val, '==': count == val, '!=': count != val}
                    result = op_map.get(op, False)
                    
                    if result:
                        self.log_execution(f"Step {self.current_step_index + 1}: Color Count SUCCEEDED. Found {count} blob(s). Condition '{expression_str}' is TRUE.", "green")
                        step['_last_run_info'] = {'timestamp': time.time(), 'result': True, 'details': f"Found {count} color blobs. Expression '{count} {op} {val}' was {result}."}
                        step['_count_current_cycle'] = 0
                        step_succeeded = True
                    else:
                        self.log_execution(f"Step {self.current_step_index + 1}: Color Count FAILED. Found {count} blob(s). Condition '{expression_str}' is FALSE.", "orange")
                        step['_last_run_info'] = {'timestamp': time.time(), 'result': False, 'details': f"Found {count} color blobs. Expression '{count} {op} {val}' was {result}."}
                        step['_count_current_cycle'] += 1
                        self.last_detection_info.set(f"Color Count: Failed. Cycle {step['_count_current_cycle']}/{max_cycles}")
                        if step['_count_current_cycle'] >= max_cycles:
                            self.log_execution(f"Step {self.current_step_index + 1}: Color Count failed after {max_cycles} cycle(s).", "orange")
                            step['_count_current_cycle'] = 0
                            self.handle_timeout()
                            return
                        else:
                            self.executor_after_id = self.root.after(int(self.scan_interval.get() * 1000), self.run_step_executor)
                            return
                except (ValueError, IndexError) as e:
                    self.log_execution(f"Step {self.current_step_index + 1}: Invalid expression '{expression_str}'. Failing: {e}", "red")
                    step['_count_current_cycle'] = 0
                    self.handle_timeout()
                    return

            elif step['type'] == 'png' and step.get('action') == 'PNG Count':
                area = step.get('area') or (self.area_x1.get(), self.area_y1.get(), self.area_x2.get(), self.area_y2.get())
                w, h = area[2] - area[0], area[3] - area[1]
                if w < 1 or h < 1:
                    self.log_execution(f"Step {self.current_step_index + 1}: Invalid area for PNG Count. Failing.", "red")
                    self.handle_timeout(); return

                screenshot = pyautogui.screenshot(region=(area[0], area[1], w, h))
                screen_cv = cv2.cvtColor(np.array(screenshot), cv2.COLOR_RGB2BGR)
                
                count = self.find_and_count_png(screen_cv, area[0:2], step)
                expression_str = step.get('count_expression', '>= 1')
                
                step.setdefault('_count_current_cycle', 0)
                max_cycles = step.get('count_max_cycles', 1)
                
                self.last_detection_info.set(f"PNG Count: Found {count}. Condition: {expression_str}")
                
                try:
                    expression = expression_str.split()
                    if len(expression) != 2: raise ValueError("Expression must have 2 parts (e.g., '>= 5')")
                    op, val = expression[0], int(expression[1])
                    op_map = {'>': count > val, '<': count < val, '>=': count >= val, '<=': count <= val, '==': count == val, '!=': count != val}
                    result = op_map.get(op, False)
                    
                    if result:
                        self.log_execution(f"Step {self.current_step_index + 1}: PNG Count SUCCEEDED. Found {count} instance(s). Condition '{expression_str}' is TRUE.", "green")
                        step['_last_run_info'] = {'timestamp': time.time(), 'result': True, 'details': f"Found {count} instances. Expression '{count} {op} {val}' was {result}."}
                        step['_count_current_cycle'] = 0
                        step_succeeded = True
                    else:
                        self.log_execution(f"Step {self.current_step_index + 1}: PNG Count FAILED. Found {count} instance(s). Condition '{expression_str}' is FALSE.", "orange")
                        step['_last_run_info'] = {'timestamp': time.time(), 'result': False, 'details': f"Found {count} instances. Expression '{count} {op} {val}' was {result}."}
                        step['_count_current_cycle'] += 1
                        self.last_detection_info.set(f"PNG Count: Failed. Cycle {step['_count_current_cycle']}/{max_cycles}")
                        if step['_count_current_cycle'] >= max_cycles:
                            self.log_execution(f"Step {self.current_step_index + 1}: PNG Count failed after {max_cycles} cycle(s).", "orange")
                            step['_count_current_cycle'] = 0
                            self.handle_timeout()
                            return
                        else:
                            self.executor_after_id = self.root.after(int(self.scan_interval.get() * 1000), self.run_step_executor)
                            return
                except (ValueError, IndexError) as e:
                    self.log_execution(f"Step {self.current_step_index + 1}: Invalid expression '{expression_str}'. Failing: {e}", "red")
                    step['_count_current_cycle'] = 0
                    self.handle_timeout()
                    return
            
            elif step['type'] == 'location': 
                target_pos, step_succeeded = step['coords'], True
                action_text = f"Press Key '{step.get('key_to_press')}'" if step.get('action') == 'Key Press' else f"{step.get('action')} at {step.get('coords')}"
                self.last_detection_info.set(f"Action: {action_text}")
                step['_last_run_info'] = {'timestamp': time.time(), 'result': True, 'details': f"Action '{step.get('action')}' scheduled."}

            elif step['type'] == 'logical':
                step_succeeded, should_return = self.execute_logical_step(step)
                if should_return: return

            else: # Threaded Detection for regular PNG, Color
                with self.detection_lock:
                    if self.detection_thread and self.detection_thread.is_alive():
                        self.executor_after_id = self.root.after(int(self.scan_interval.get() * 1000), self.run_step_executor); return
                    
                    if self.detection_result:
                        result_type, target_pos, confidence = self.detection_result
                        current_step_type = step.get('logical_type') or step.get('type')
                        
                        if result_type == current_step_type:
                            if result_type == 'png':
                                if target_pos:
                                    step_succeeded = True
                                    self.last_detection_info.set(f"PNG Found: {confidence*100:.1f}%")
                                    self.log_execution(f"Step {self.current_step_index + 1}: PNG FOUND at {target_pos} with {confidence*100:.1f}% confidence.", "green")
                                    step['_last_run_info'] = {'timestamp': time.time(), 'result': True, 'details': f"Found at {target_pos} with {confidence*100:.1f}% confidence."}
                            
                            elif result_type == 'color':
                                contour_area = confidence # In color detection, confidence holds the area
                                if target_pos:
                                    step_succeeded = True
                                    self.last_detection_info.set(f"Color Found: Area {contour_area:.0f}px")
                                    if step.get('pixel_detect_enabled'):
                                        self.log_execution(f"Step {self.current_step_index + 1}: Pixel Color FOUND at {target_pos}.", "green")
                                        step['_last_run_info'] = {'timestamp': time.time(), 'result': True, 'details': f"Found pixel at {target_pos}."}
                                    else:
                                        self.log_execution(f"Step {self.current_step_index + 1}: Color Area FOUND at {target_pos} with area {contour_area:.0f}px.", "green")
                                        step['_last_run_info'] = {'timestamp': time.time(), 'result': True, 'details': f"Found at {target_pos} with area {contour_area:.0f}px."}
                        
                        self.detection_result = None

                if not step_succeeded:
                    area = step.get('area') or (self.area_x1.get(),self.area_y1.get(),self.area_x2.get(),self.area_y2.get())
                    w,h = area[2]-area[0],area[3]-area[1]
                    if w < 1 or h < 1: 
                        self.executor_after_id = self.root.after(int(self.scan_interval.get()*1000),self.run_step_executor); return
                    
                    screenshot = pyautogui.screenshot(region=(area[0],area[1],w,h))
                    screen_cv = cv2.cvtColor(np.array(screenshot),cv2.COLOR_RGB2BGR)

                    if step['type'] == 'png':
                        self.last_detection_info.set(f"PNG: Searching for {os.path.basename(step.get('path'))}...")
                        self.log_execution(f"Step {self.current_step_index + 1}: Searching for PNG '{os.path.basename(step.get('path'))}' in area {area} (Thresh: {step.get('threshold')}).")
                        self.detection_thread = threading.Thread(target=self._perform_png_detection_in_thread, args=(screen_cv, area[0:2], step), daemon=True)
                    elif step['type'] == 'color':
                        if step.get('pixel_detect_enabled'):
                            self.last_detection_info.set(f"Color: Searching for RGB {step.get('rgb')} at pixel {step.get('pixel_coords')}...")
                            self.log_execution(f"Step {self.current_step_index + 1}: Checking for Color {step.get('rgb')} at pixel {step.get('pixel_coords')} (Tol: {step.get('tolerance')}, Space: {step.get('color_space')}).")
                        else:
                            self.last_detection_info.set(f"Color: Searching for RGB {step.get('rgb')}...")
                            self.log_execution(f"Step {self.current_step_index + 1}: Searching for Color {step.get('rgb')} in area {area} (Tol: {step.get('tolerance')}, Space: {step.get('color_space')}).")
                        self.detection_thread = threading.Thread(target=self._perform_color_detection_in_thread, args=(screen_cv, area[0:2], step), daemon=True)
                    
                    if self.detection_thread: self.detection_thread.start()
                    self.executor_after_id = self.root.after(int(self.scan_interval.get()*1000), self.run_step_executor)
                    return

            # --- ACTION AND FLOW CONTROL (After a step succeeds) ---
            if step_succeeded:
                if self.timeout_countdown_id:
                    self.root.after_cancel(self.timeout_countdown_id)
                    self.timeout_countdown_id = None
                self.timeout_countdown_label.config(text="")
                
                if step['type'] == 'location':
                    if step.get('action') == 'Key Press':
                        pyautogui.press(step.get('key_to_press'))
                        self.log_execution(f"Step {self.current_step_index + 1}: Pressed key '{step.get('key_to_press')}'.")
                    else: 
                        self.execute_action_on_pos(step.get('action'), target_pos)
                elif step['type'] != 'logical': # For regular PNG and Color
                    self.execute_action_on_pos(step.get('action'), target_pos)
                
                self.handle_flow_control('on_success_action', 'on_success_goto_step')
            else:
                 self.executor_after_id = self.root.after(int(self.scan_interval.get()*1000), self.run_step_executor)

        except Exception as e:
            messagebox.showerror("Execution Error", str(e)); self.log(f"Execution Error: {e}", "red"); self.stop("Status: Stopped due to error", color_state='red'); return

    def execute_logical_step(self, step):
        logical_type = step.get('logical_type')
        if logical_type == 'Count':
            current_val = step.get('counter_value', 0)
            max_count = step.get('max_count', 0)
            self.last_detection_info.set(f"Count: {current_val + 1} / {max_count if max_count > 0 else '∞'}")
            step['counter_value'] = current_val + 1
            step['_last_run_info'] = {'timestamp': time.time(), 'result': True, 'details': f"Counter incremented to {step['counter_value']}."}
            self.log_execution(f"Step {self.current_step_index + 1}: Count is now {step['counter_value']}/{max_count if max_count > 0 else '∞'}.")
            
            if max_count > 0 and step['counter_value'] >= max_count:
                step['_last_run_info']['result'] = 'Reached'
                step['_last_run_info']['details'] += f" Max count of {max_count} reached."
                self.log_execution(f"Step {self.current_step_index + 1}: Max count of {max_count} reached.", "orange")
                self.handle_flow_control('on_count_reached_action', 'on_count_reached_goto_step')
                if step.get('reset_on_reach', False):
                    step['counter_value'] = 0
                    self.log_execution(f"Step {self.current_step_index + 1}: Counter reset after reaching max count.")
                return False, True 
            return True, False

        elif logical_type == 'Wait':
            if step.get('timer_start_time') is None:
                step['timer_start_time'] = time.time()
            elapsed = time.time() - step['timer_start_time']
            max_time = step.get('max_time', 0)
            self.last_detection_info.set(f"Wait: {elapsed:.1f}s / {max_time:.1f}s")
            step['_last_run_info'] = {'timestamp': time.time(), 'result': 'Waiting', 'details': f"Elapsed: {elapsed:.1f}s"}
            if elapsed >= max_time:
                step['_last_run_info']['result'] = True
                step['_last_run_info']['details'] = f"Waited for {elapsed:.2f}s."
                self.log_execution(f"Step {self.current_step_index + 1}: Wait timer of {max_time}s finished.")
                step['last_cycle_time'] = round(elapsed, 2); step['timer_start_time'] = None
                return True, False
            else:
                self.executor_after_id = self.root.after(int(self.scan_interval.get()*1000), self.run_step_executor)
                return False, True

        elif logical_type == 'Type Text':
            text_to_type = ""
            source = step.get('text_source', 'Static Text')
            if source == 'GE Interface':
                field = step.get('ge_data_field')
                if field == "Item Name":
                    text_to_type = self.ge_interface_item_name.get()
                elif field == "Quantity":
                    text_to_type = self.ge_interface_item_quantity.get()
                else:
                    if self.ge_interface_last_data:
                        try:
                            data = self.ge_interface_last_data
                            high_price = int(data.get('high', 0))
                            low_price = int(data.get('low', 0))
                            quantity = int(self.ge_interface_item_quantity.get())
                            
                            if field == "Calculated Buy Price":
                                text_to_type = self._calculate_ge_price('buy', high_price, low_price)
                            elif field == "Calculated Sell Price":
                                text_to_type = self._calculate_ge_price('sell', high_price, low_price)
                            elif field == "Calculated Buy Total":
                                price = self._calculate_ge_price('buy', high_price, low_price)
                                text_to_type = price * quantity
                            elif field == "Calculated Sell Total":
                                price = self._calculate_ge_price('sell', high_price, low_price)
                                text_to_type = price * quantity
                            
                        except (ValueError, TypeError) as e:
                            self.log_execution(f"Step {self.current_step_index + 1}: Error processing GE data: {e}", "red")
                            text_to_type = "ERROR"
                    else:
                        self.log_execution(f"Step {self.current_step_index + 1}: No GE data available to type. Fetch data first.", "orange")
                        text_to_type = ""
            else: 
                text_to_type = step.get('text_to_type', '')
            
            self.last_detection_info.set(f"Type Text: Typing '{str(text_to_type)[:25]}...'")
            pyautogui.write(str(text_to_type).replace(',', ''), interval=0.05)
            if step.get('press_enter', False):
                delay = step.get('enter_press_delay', 0.1)
                time.sleep(delay)
                pyautogui.press('enter')
            
            step['_last_run_info'] = {'timestamp': time.time(), 'result': True, 'details': f"Typed '{text_to_type}' from {source}."}
            self.log_execution(f"Step {self.current_step_index + 1}: Typed text '{text_to_type}' from {source}.")
            return True, False

        elif logical_type == 'GE Inject':
            field = step.get('ge_inject_field', 'Name')
            if field == 'Quantity':
                quantity_to_inject = step.get('ge_inject_quantity', '1')
                self.ge_interface_item_quantity.set(quantity_to_inject)
                self.last_detection_info.set(f"GE Inject: Set quantity to '{quantity_to_inject}'")
                step['_last_run_info'] = {'timestamp': time.time(), 'result': True, 'details': f"Injected quantity '{quantity_to_inject}'."}
                self.log_execution(f"Step {self.current_step_index + 1}: Injected quantity '{quantity_to_inject}' into GE Interface.")
            else: # Default to Name
                item_name = step.get('ge_inject_name', '')
                self.ge_interface_item_name.set(item_name)
                self.last_detection_info.set(f"GE Inject: Set item to '{item_name}'")
                step['_last_run_info'] = {'timestamp': time.time(), 'result': True, 'details': f"Injected item name '{item_name}'."}
                self.log_execution(f"Step {self.current_step_index + 1}: Injected item name '{item_name}' into GE Interface.")
            
            if step.get('ge_inject_refresh', False):
                self.log_execution(f"Step {self.current_step_index + 1}: Triggering GE price refresh after inject.")
                self.update_ge_interface_price()

            return True, False
        
        elif logical_type == 'Settings Inject':
            setting_name = step.get('inject_setting_name')
            new_value_str = step.get('inject_setting_value')
            
            setting_map = {
                "Location Offset (±px)": {'model': self.loc_offset_variance, 'type': int},
                "Speed Variance (±s)": {'model': self.speed_variance, 'type': float},
                "Hold Variance (±s)": {'model': self.hold_duration_variance, 'type': float},
                "Scan Interval (s)": {'model': self.scan_interval, 'type': float},
                "Base Hold Duration (s)": {'model': self.hold_duration, 'type': float}
            }

            if setting_name in setting_map:
                try:
                    setting_info = setting_map[setting_name]
                    converted_value = setting_info['type'](new_value_str)
                    setting_info['model'].set(converted_value)
                    self._sync_global_settings_ui_from_model() # Update UI
                    
                    self.last_detection_info.set(f"Inject: Set {setting_name} to {converted_value}")
                    details = f"Injected '{setting_name}' = {converted_value}"
                    step['_last_run_info'] = {'timestamp': time.time(), 'result': True, 'details': details}
                    self.log_execution(f"Step {self.current_step_index + 1}: {details}.")
                    return True, False
                except (ValueError, tk.TclError) as e:
                    details = f"Invalid value '{new_value_str}' for {setting_name}. Error: {e}"
                    step['_last_run_info'] = {'timestamp': time.time(), 'result': False, 'details': details}
                    self.log_execution(f"Step {self.current_step_index + 1}: {details}", "red")
            else:
                details = f"Unknown setting '{setting_name}' to inject."
                step['_last_run_info'] = {'timestamp': time.time(), 'result': False, 'details': details}
                self.log_execution(f"Step {self.current_step_index + 1}: {details}", "red")

            return False, False

        elif logical_type == 'Movement Detect':
            area = step.get('area') or (self.area_x1.get(), self.area_y1.get(), self.area_x2.get(), self.area_y2.get())
            w, h = area[2] - area[0], area[3] - area[1]
            if w < 1 or h < 1:
                self.log_execution(f"Step {self.current_step_index + 1}: Invalid area for Movement Detect. Failing.", "red")
                self.handle_timeout()
                return False, True

            screenshot = pyautogui.screenshot(region=(area[0], area[1], w, h))
            current_frame_cv = cv2.cvtColor(np.array(screenshot), cv2.COLOR_RGB2GRAY)
            
            previous_frame = step.get('_previous_frame_for_movement')

            if previous_frame is None:
                step['_previous_frame_for_movement'] = current_frame_cv
                self.last_detection_info.set("Movement: 1st frame captured. Waiting for 2nd...")
                self.log_execution(f"Step {self.current_step_index + 1}: Captured first frame for movement comparison.")
                step['_last_run_info'] = {'timestamp': time.time(), 'result': 'Waiting', 'details': 'First frame captured.'}
                self.executor_after_id = self.root.after(int(self.scan_interval.get() * 1000), self.run_step_executor)
                return False, True
            else:
                if previous_frame.shape != current_frame_cv.shape:
                    self.log_execution(f"Step {self.current_step_index + 1}: Frame dimension mismatch. Resetting comparison.", "orange")
                    step['_previous_frame_for_movement'] = current_frame_cv
                    self.executor_after_id = self.root.after(int(self.scan_interval.get() * 1000), self.run_step_executor)
                    return False, True

                diff = cv2.absdiff(previous_frame, current_frame_cv)
                _, thresholded_diff = cv2.threshold(diff, 30, 255, cv2.THRESH_BINARY)
                
                non_zero_count = np.count_nonzero(thresholded_diff)
                total_pixels = thresholded_diff.size
                change_percentage = (non_zero_count / total_pixels) * 100 if total_pixels > 0 else 0
                
                tolerance = step.get('movement_tolerance', 5.0)
                
                self.last_detection_info.set(f"Movement: {change_percentage:.2f}% changed (Tolerance: {tolerance}%)")
                
                step['_previous_frame_for_movement'] = None

                if change_percentage <= tolerance:
                    self.log_execution(f"Step {self.current_step_index + 1}: Stillness detected between cycles ({change_percentage:.2f}% <= {tolerance}%). Success.")
                    step['_last_run_info'] = {'timestamp': time.time(), 'result': True, 'details': f"Stillness detected. Change: {change_percentage:.2f}%."}
                    return True, False
                else:
                    self.log_execution(f"Step {self.current_step_index + 1}: Movement detected between cycles ({change_percentage:.2f}%). Continuing.")
                    step['_last_run_info'] = {'timestamp': time.time(), 'result': 'Waiting', 'details': f"Movement ongoing. Change: {change_percentage:.2f}%."}
                    self.executor_after_id = self.root.after(int(self.scan_interval.get() * 1000), self.run_step_executor)
                    return False, True
        
        elif logical_type == 'Number':
            area = step.get('area') or (self.area_x1.get(), self.area_y1.get(), self.area_x2.get(), self.area_y2.get())
            expression_str = step.get('expression', '> 0')
            w, h = area[2] - area[0], area[3] - area[1]
            if w < 1 or h < 1:
                return False, False 
            
            self.log_execution(f"Step {self.current_step_index + 1}: Performing OCR in area {area} with expression '{expression_str}'.")
            try:
                screenshot = pyautogui.screenshot(region=(area[0], area[1], w, h))
                screen_cv = cv2.cvtColor(np.array(screenshot), cv2.COLOR_RGB2BGR)

                image_mode = step.get('image_mode', 'Grayscale')
                gray = cv2.cvtColor(screen_cv, cv2.COLOR_BGR2GRAY)

                if image_mode == 'Binary (B&W)':
                    inverted = cv2.bitwise_not(gray)
                    _, processed_for_ocr = cv2.threshold(inverted, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
                elif image_mode == 'Grayscale':
                    processed_for_ocr = cv2.bitwise_not(gray)
                else: 
                    processed_for_ocr = cv2.cvtColor(screen_cv, cv2.COLOR_BGR2RGB)

                psm_mode = self.psm_options.get(step.get('psm_mode'), '6')
                oem_mode = self.oem_options.get(step.get('oem_mode'), '3')
                ocr_config = f'--oem {oem_mode} --psm {psm_mode} -c tessedit_char_whitelist=0123456789:;,.-'
                ocr_text = pytesseract.image_to_string(Image.fromarray(processed_for_ocr), config=ocr_config)
                cleaned_text = "".join(filter(lambda x: x in '0123456789.-', ocr_text))

                self.log_execution(f" > OCR Raw Text: '{ocr_text.strip()}'. Cleaned Number: '{cleaned_text}'.")
                
                if not cleaned_text:
                    self.last_detection_info.set("OCR: No number detected in area.")
                    self.log_execution(" > OCR FAILED: No valid number characters found in area.", "orange")
                    return False, False 

                num = float(cleaned_text)
                expression = expression_str.split()
                if len(expression) != 2: raise ValueError("Invalid expression format")
                
                op, val = expression[0], float(expression[1])
                op_map = {'>': num > val, '<': num < val, '>=': num >= val, '<=': num <= val, '==': num == val, '!=': num != val}
                result = op_map.get(op, False)

                self.last_detection_info.set(f"OCR: '{num}'. Condition met: {result}")
                step['_last_run_info'] = {'timestamp': time.time(), 'result': result, 'details': f"OCR found '{num}'. Condition success: {result}."}
                
                if result:
                    self.log_execution(f" > Evaluation: '{num} {op} {val}' is TRUE. SUCCEEDED.", "green")
                else:
                    self.log_execution(f" > Evaluation: '{num} {op} {val}' is FALSE. FAILED.", "orange")

                return result, False

            except Exception as e:
                self.last_detection_info.set(f"OCR Error: Retrying...")
                self.log_execution(f" > OCR ERROR: {e}", "red")
                print(f"Error during OCR in step {self.current_step_index + 1}: {e}")
                return False, False

        return False, False

    def handle_timeout(self):
        step = self.steps[self.current_step_index]
        
        is_count_fail = (step.get('type') == 'png' and step.get('action') == 'PNG Count') or \
                        (step.get('type') == 'color' and step.get('action') == 'Color Count')
        is_number_fail = step.get('logical_type') == 'Number'
        
        if is_count_fail or is_number_fail:
            log_msg = f"Step {self.current_step_index+1} failed."
        else:
            log_msg = f"Step {self.current_step_index+1} timed out."

        self.log_execution(log_msg, "orange")
        
        step['_last_run_info']['result'] = 'Timeout'
        step['_last_run_info']['details'] = f"Step failed or timed out after {step.get('timeout', 0)}s."

        action = step['on_timeout_action']
        if action == 'Stop': self.stop(f"Status: Stopped on timeout at Step {self.current_step_index + 1}", color_state='orange'); return
        
        next_index = self.current_step_index + 1 if action == 'Next Step' else step['on_timeout_goto_step'] - 1
        self.current_step_index = next_index
        
        # Schedule advance_step to break any potential recursion loops.
        self.executor_after_id = self.root.after(1, self.advance_step)

    def handle_flow_control(self, action_key, goto_key):
        step = self.steps[self.current_step_index]; action = step.get(action_key, 'Stop')
        if action == 'Stop': self.stop(f"Status: Stopped by flow control at Step {self.current_step_index + 1}", color_state='orange'); return
        next_index = self.current_step_index + 1 if action == 'Next Step' else step.get(goto_key, 1) - 1
        self.current_step_index = next_index
        delay_key = 'delay_after'
        if action_key == 'on_count_reached_action': delay_key = 'on_count_reached_delay'
        self.start_delay_countdown(step.get(delay_key, 0))

    def start_delay_countdown(self, delay_seconds, next_action_func=None):
        if self.delay_countdown_id: self.root.after_cancel(self.delay_countdown_id)
        if next_action_func is None: next_action_func = self.advance_step
        if delay_seconds > 0: end_time = time.time() + delay_seconds; self.update_delay_countdown(end_time, next_action_func)
        else: next_action_func()

    def update_delay_countdown(self, end_time, next_action_func):
        remaining = end_time - time.time()
        if remaining > 0 and self.running: self.delay_countdown_label.config(text=f"Next step in {remaining:.1f}s..."); self.delay_countdown_id = self.root.after(100, lambda: self.update_delay_countdown(end_time, next_action_func))
        elif self.running: self.delay_countdown_label.config(text=""); next_action_func()

    def update_timeout_countdown(self, start_time, total_timeout):
        if not self.running: return
        remaining = total_timeout - (time.time() - start_time)
        if remaining > 0: self.timeout_countdown_label.config(text=f"Timeout in {remaining:.1f}s..."); self.timeout_countdown_id = self.root.after(100, lambda: self.update_timeout_countdown(start_time, total_timeout))
        else: self.timeout_countdown_label.config(text="")

