import tkinter as tk
from tkinter import filedialog, messagebox
import json
import copy
import os

class FileOpsMixin:
    def add_step(self, step_type):
        new_x, new_y = 50, 50
        if self.steps: last_step = self.steps[-1]; self._calculate_node_size(len(self.steps)-1); new_x, new_y = last_step.get('x', 50), last_step.get('y', 50) + last_step.get('_height', 60)/self.zoom_factor + 40
        num_steps = len(self.steps)
        step_name_map = {'location': 'New Click / Press'}
        step_name = step_name_map.get(step_type, f'New {step_type.title()}')
        step_defaults = {
            'type': step_type, 'name': step_name, 'delay_after': 1.0, 'timeout': 0, 
            'on_timeout_action': 'Stop', 'on_timeout_goto_step': 1, 'on_success_action': 'Next Step', 
            'on_success_goto_step': num_steps + 2, 'x': new_x, 'y': new_y, 'enable_logging': True, 'show_area': False,
            '_last_run_info': {'timestamp': None, 'result': None, 'details': 'Not yet run'}
        }
        if step_type == 'color': 
            step_defaults.update({
                'action':'Click Object', 'rgb':(255,0,0), 'tolerance':2, 'area':None, 'color_space': 'HSV',
                'min_pixel_area': 10,
                'count_expression': '>= 1',
                'count_max_cycles': 1,
            })
        elif step_type == 'png': 
            step_defaults.update({
                'action':'Click Object', 'mode':'file', 'path':'', 'threshold':0.8, 'area':None, 
                'image_mode': 'Grayscale', 'find_first_match': True,
                'count_expression': '>= 1',
                'count_max_cycles': 1
            })
        elif step_type == 'location': step_defaults.update({'action': 'Left Click', 'coords':(100,100), 'key_to_press': ''})
        elif step_type == 'logical':
            step_defaults.update({
                'action': 'Execute', 'logical_type': 'Count', 'counter_value': 0, 'max_count': 0, 'reset_on_start': False, 'reset_on_reach': False,
                'delay_after': 0.1, 'timer_start_time': None, 'last_cycle_time': 'N/A', 'max_time': 5,
                'text_to_type': '', 'press_enter': False, 'enter_press_delay': 0.1, 'text_source': 'Static Text', 'ge_data_field': 'Calculated Buy Price',
                'ge_inject_name': '', 'ge_inject_field': 'Name', 'ge_inject_quantity': '1',
                'ge_inject_refresh': False,
                'inject_setting_name': 'Location Offset (±px)', 'inject_setting_value': '4',
                'on_count_reached_action': 'Stop', 'on_count_reached_goto_step': 1, 'on_count_reached_delay': 1.0,
                'expression': '> 0', 'area': None, 'timeout': 5, 'on_timeout_action': 'Next Step', 
                'image_mode': 'Grayscale', 'psm_mode': '6: Assume a single uniform block of text.', 'oem_mode': '3: Default, based on what is available.',
                'movement_tolerance': 5.0,
                '_previous_frame_for_movement': None
            })
        self.steps.append(step_defaults); self.log(f"Added Step {len(self.steps)}: {step_name}"); 
        self.selected_items = [{'type': 'step', 'index': len(self.steps)-1}]
        self.populate_properties_panel()
        self.redraw_flowchart()

    def add_annotation(self):
        new_note = {'type': 'note', 'x': 60, 'y': 60, 'width': 200, 'height': 120, 'text': 'New Note', 'color': '#fffacd', 'opacity': '0% (Border Only)'}
        self.annotations.append(new_note); self.log("Added a new note to the flowchart."); 
        self.selected_items = [{'type': 'note', 'index': len(self.annotations)-1}]
        self.populate_properties_panel()
        self.redraw_flowchart()

    def remove_step(self):
        if not (self.selected_items and len(self.selected_items) == 1 and self.selected_items[0]['type'] == 'step'): return
        index_to_remove = self.selected_items[0]['index']
        if not messagebox.askyesno("Confirm Delete", f"Are you sure you want to delete Step {index_to_remove + 1}?"): return
        
        if index_to_remove in self.area_overlays:
            self.area_overlays[index_to_remove].destroy()
            del self.area_overlays[index_to_remove]

        self.steps.pop(index_to_remove)
        for step in self.steps:
            for goto_key in ['on_success_goto_step', 'on_timeout_goto_step', 'on_count_reached_goto_step']:
                if goto_key not in step: continue
                current_goto = step.get(goto_key, 0)
                if current_goto > index_to_remove + 1: step[goto_key] -= 1
                elif current_goto == index_to_remove + 1: step[goto_key] = 1
        self.log(f"Removed step {index_to_remove + 1}.")
        self.selected_items = []
        self.populate_properties_panel()
        self.redraw_flowchart()

    def remove_note(self):
        if not (self.selected_items and len(self.selected_items) == 1 and self.selected_items[0]['type'] == 'note'): return
        index_to_remove = self.selected_items[0]['index']
        if not messagebox.askyesno("Confirm Delete", f"Are you sure you want to delete this note?"): return
        self.annotations.pop(index_to_remove); self.log(f"Removed note."); 
        self.selected_items = []
        self.populate_properties_panel()
        self.redraw_flowchart()

    def duplicate_step(self):
        if not (self.selected_items and len(self.selected_items) == 1 and self.selected_items[0]['type'] == 'step'): return
        index = self.selected_items[0]['index']; original_step = self.steps[index]; new_step = copy.deepcopy(original_step); self._calculate_node_size(index)
        new_step['x'] = original_step.get('x', 50) + 20; new_step['y'] = original_step.get('y', 50) + original_step.get('_height', 60)/self.zoom_factor + 30; new_step['name'] = original_step.get('name', 'Unnamed') + " (Copy)"
        new_index = len(self.steps)
        if new_step.get('on_success_goto_step') == index + 1: new_step['on_success_goto_step'] = new_index + 1
        if new_step.get('on_timeout_goto_step') == index + 1: new_step['on_timeout_goto_step'] = new_index + 1
        self.steps.append(new_step); self.log(f"Duplicated Step {index + 1} to new Step {len(self.steps)}."); 
        self.selected_items = [{'type': 'step', 'index': len(self.steps) - 1}]
        self.populate_properties_panel()
        self.redraw_flowchart()

    def duplicate_note(self):
        if not (self.selected_items and len(self.selected_items) == 1 and self.selected_items[0]['type'] == 'note'): return
        index = self.selected_items[0]['index']
        original_note = self.annotations[index]
        new_note = copy.deepcopy(original_note)
        new_note['x'] = original_note.get('x', 50) + 20
        new_note['y'] = original_note.get('y', 50) + 20
        self.annotations.append(new_note)
        self.log("Duplicated note.")
        self.selected_items = [{'type': 'note', 'index': len(self.annotations) - 1}]
        self.populate_properties_panel()
        self.redraw_flowchart()

    def delete_selected(self):
        if not self.selected_items or self.selected_items[0]['type'] != 'step':
            return

        remove_set = {item['index'] for item in self.selected_items}

        for index in remove_set:
            if index in self.area_overlays:
                self.area_overlays[index].destroy()
                del self.area_overlays[index]

        # Build old→new index map in a single pass (no shifting during iteration)
        index_map = {}
        new_idx = 0
        for old_idx in range(len(self.steps)):
            if old_idx not in remove_set:
                index_map[old_idx] = new_idx
                new_idx += 1

        for index in sorted(remove_set, reverse=True):
            self.steps.pop(index)

        for step in self.steps:
            for goto_key in ['on_success_goto_step', 'on_timeout_goto_step', 'on_count_reached_goto_step']:
                if goto_key not in step:
                    continue
                old_target = step.get(goto_key, 1) - 1
                step[goto_key] = index_map.get(old_target, 0) + 1

        self.log(f"Removed {len(remove_set)} steps.")
        self.selected_items = []
        self.populate_properties_panel()
        self.redraw_flowchart()
        self.update_all_area_overlays()

    def duplicate_selected(self):
        if not self.selected_items or self.selected_items[0]['type'] != 'step': return
        original_indices = sorted([item['index'] for item in self.selected_items]);
        if not original_indices: return
        new_steps, old_to_new_index_map, current_len = [], {}, len(self.steps)
        for i, original_index in enumerate(original_indices):
            new_index = current_len + i; old_to_new_index_map[original_index] = new_index
            original_step = self.steps[original_index]; new_step = copy.deepcopy(original_step); new_step['name'] += " (Copy)"; new_steps.append(new_step)
        for i, new_step in enumerate(new_steps):
            original_index = original_indices[i]
            for action_key, goto_key in [('on_success_action', 'on_success_goto_step'), ('on_timeout_action', 'on_timeout_goto_step'), ('on_count_reached_action', 'on_count_reached_goto_step')]:
                if action_key not in new_step: continue
                action = new_step.get(action_key)
                if action == 'Go to Step':
                    target_original_index = new_step.get(goto_key, 1) - 1
                    if target_original_index in old_to_new_index_map: new_step[goto_key] = old_to_new_index_map[target_original_index] + 1
                    else: new_step[action_key] = 'Stop'
                elif action == 'Next Step':
                    next_original_index = original_index + 1
                    if next_original_index not in old_to_new_index_map: new_step[action_key] = 'Stop'
        min_x = min(self.steps[i].get('x', 0) for i in original_indices); max_y = 0
        for i in original_indices:
            self._calculate_node_size(i)
            current_max_y = self.steps[i].get('y', 0) + self.steps[i].get('_height', 60) / self.zoom_factor
            if current_max_y > max_y: max_y = current_max_y
        y_offset = max_y + 60
        for i, new_step in enumerate(new_steps):
            original_pos_x = self.steps[original_indices[i]].get('x', 0); original_pos_y = self.steps[original_indices[i]].get('y', 0)
            new_step['x'] = original_pos_x ; new_step['y'] = y_offset + (original_pos_y - self.steps[original_indices[0]].get('y', 0))
        self.steps.extend(new_steps); self.log(f"Duplicated {len(new_steps)} steps.")
        new_indices_range = range(len(self.steps) - len(new_steps), len(self.steps))
        self.selected_items = [{'type': 'step', 'index': i} for i in new_indices_range]
        self.populate_properties_panel(); self.redraw_flowchart()
        self.update_all_area_overlays() # Redraw overlays for new duplicated steps

    def reset_logical_counter(self):
        if not (self.selected_items and len(self.selected_items) == 1 and self.selected_items[0]['type'] == 'step'): return
        step = self.steps[self.selected_items[0]['index']]
        if step.get('logical_type') == 'Count': step['counter_value'] = 0; self.log(f"Reset counter for Step {self.selected_items[0]['index'] + 1}."); self.populate_properties_panel()
       
    def reset_timer(self):
        if not (self.selected_items and len(self.selected_items) == 1 and self.selected_items[0]['type'] == 'step'): return
        step = self.steps[self.selected_items[0]['index']]
        if step.get('logical_type') == 'Wait': step['timer_start_time'] = None; step['last_cycle_time'] = 'N/A'; self.log(f"Reset wait timer for Step {self.selected_items[0]['index'] + 1}."); self.populate_properties_panel()

    def export_to_json(self):
        filepath = filedialog.asksaveasfilename(defaultextension=".json", filetypes=[("JSON Files", "*.json")], title="Export Flowchart")
        if not filepath: return
        steps_to_save = copy.deepcopy(self.steps)
        for s in steps_to_save:
            s.pop('_width', None)
            s.pop('_height', None)
            s.pop('_last_run_info', None)
            s.pop('_previous_frame_for_movement', None)
        settings = {
            "global_settings": {
                "mouse_move_mode": self.mouse_move_mode.get(),
                "mouse_speed": self.mouse_speed.get(),
                "pixels_per_second": self.pixels_per_second.get(),
                "min_move_time": self.min_move_time.get(),
                "max_move_time": self.max_move_time.get(),
                "scan_interval": self.scan_interval.get(), 
                "hold_duration": self.hold_duration.get(), 
                "loc_offset_variance": self.loc_offset_variance.get(), 
                "speed_variance": self.speed_variance.get(), 
                "hold_duration_variance": self.hold_duration_variance.get(), 
                "area_x1": self.area_x1.get(), "area_y1": self.area_y1.get(), "area_x2": self.area_x2.get(), 
                "area_y2": self.area_y2.get(), "hide_on_select": self.hide_on_select.get(),
                "start_at_stopped_pos": self.start_at_stopped_pos.get(),
                # --- Grid Settings ---
                "grid_visible": self.grid_visible.get(),
                "grid_latching": self.grid_latching.get(),
                "grid_spacing": self.grid_spacing.get(),
                "grid_opacity": self.grid_opacity.get()
            }, 
            "steps": steps_to_save, 
            "annotations": self.annotations 
        }
        
        ge_step_exists = any(s.get('logical_type') == 'GE Inject' or s.get('text_source') == 'GE Interface' for s in steps_to_save)
        if ge_step_exists:
            settings["ge_interface_settings"] = {
                'item_name': self.ge_interface_item_name.get(),
                'quantity': self.ge_interface_item_quantity.get(),
                'buy_price_strategy': self.ge_interface_buy_price_strategy.get(),
                'sell_price_strategy': self.ge_interface_sell_price_strategy.get(),
                'buy_custom_price': self.ge_interface_buy_custom_price.get(),
                'sell_custom_price': self.ge_interface_sell_custom_price.get(),
                'buy_price_margin': self.ge_interface_buy_price_margin.get(),
                'sell_price_margin': self.ge_interface_sell_price_margin.get()
            }
            self.log("GE Interface settings included in export.")
            
        try:
            with open(filepath,'w') as f: json.dump(settings,f,indent=4)
            self.log(f"Successfully exported flowchart to {os.path.basename(filepath)}.")
        except Exception as e: messagebox.showerror("Export Error",f"Failed to save file: {e}")

    def import_from_json(self):
        self.destroy_all_overlays()
        filepath = filedialog.askopenfilename(filetypes=[("JSON Files","*.json")], title="Import and Append Flowchart");
        if not filepath: return
        try:
            with open(filepath, 'r') as f: loaded_data = json.load(f)
            if "global_settings" in loaded_data:
                gs = loaded_data["global_settings"]
                
                # Handle backward compatibility for mouse mode
                if "mouse_move_mode" in gs:
                    self.mouse_move_mode.set(gs.get("mouse_move_mode", "Regular"))
                elif gs.get("enable_dynamic_speed", False):
                    self.mouse_move_mode.set("Dynamic")
                else:
                    self.mouse_move_mode.set("Regular")

                self.mouse_speed.set(gs.get("mouse_speed", 0.25))
                self.pixels_per_second.set(gs.get("pixels_per_second", 1000))
                self.min_move_time.set(gs.get("min_move_time", 0.05))
                self.max_move_time.set(gs.get("max_move_time", 0.3))
                self.scan_interval.set(gs.get("scan_interval", 0.25))
                self.hold_duration.set(gs.get("hold_duration", 0.08))
                self.loc_offset_variance.set(gs.get("loc_offset_variance", 4))
                self.speed_variance.set(gs.get("speed_variance", 0.06))
                self.hold_duration_variance.set(gs.get("hold_duration_variance", 0.03))
                self.area_x1.set(gs.get("area_x1", 0)); self.area_y1.set(gs.get("area_y1", 0))
                screen_w, screen_h = pyautogui.size()
                self.area_x2.set(gs.get("area_x2", screen_w)); self.area_y2.set(gs.get("area_y2", screen_h)); self.hide_on_select.set(gs.get("hide_on_select", True))
                self.start_at_stopped_pos.set(gs.get("start_at_stopped_pos", False))
                # --- Grid Settings ---
                self.grid_visible.set(gs.get("grid_visible", False))
                self.grid_latching.set(gs.get("grid_latching", False))
                self.grid_spacing.set(gs.get("grid_spacing", 30))
                self.grid_opacity.set(gs.get("grid_opacity", 0.3))
                self._sync_global_settings_ui_from_model(); self.apply_theme(); self.log("Loaded global settings from file.")

            if "ge_interface_settings" in loaded_data:
                gis = loaded_data["ge_interface_settings"]
                self.ge_interface_item_name.set(gis.get('item_name', ''))
                self.ge_interface_item_quantity.set(gis.get('quantity', '1'))
                self.ge_interface_buy_price_strategy.set(gis.get('buy_price_strategy', 'Flip-Buy (use Insta-Sell)'))
                self.ge_interface_sell_price_strategy.set(gis.get('sell_price_strategy', 'Flip-Sell (use Insta-Buy)'))
                self.ge_interface_buy_custom_price.set(gis.get('buy_custom_price', '0'))
                self.ge_interface_sell_custom_price.set(gis.get('sell_custom_price', '0'))
                self.ge_interface_buy_price_margin.set(gis.get('buy_price_margin', '1'))
                self.ge_interface_sell_price_margin.set(gis.get('sell_price_margin', '1'))
                self.log("Loaded GE Interface settings from file.")
                # --- FIX: Manually trigger UI update for margin widgets ---
                self._toggle_ge_buy_options()
                self._toggle_ge_sell_options()

            steps, notes = loaded_data.get("steps", []), loaded_data.get("annotations", [])
            
            cleaned_steps = []
            for s in steps:
                if s.get('type') == 'ge':
                    self.log(f"Note: Old 'GE Step' ({s.get('name')}) was ignored during import.", "orange")
                    continue
                if s.get('type') == 'number': s['type'] = 'logical'; s['logical_type'] = 'Number'
                if s.get('type') == 'location' and s.get('action') == 'Click Object': s['action'] = 'Left Click'
                if s.get('logical_type') == 'Timer': s['logical_type'] = 'Wait'
                if s.get('logical_type') == 'Number':
                    if s.get('psm_mode') and s['psm_mode'].isdigit(): s['psm_mode'] = next((k for k, v in self.psm_options.items() if v == s['psm_mode']), "6: Assume a single uniform block of text.")
                    if s.get('oem_mode') and s['oem_mode'].isdigit(): s['oem_mode'] = next((k for k, v in self.oem_options.items() if v == s['oem_mode']), "3: Default, based on what is available.")
                s['_last_run_info'] = {'timestamp': None, 'result': None, 'details': 'Imported, not yet run'}
                cleaned_steps.append(s)

            if not cleaned_steps and not notes: self.log("Imported file contains no compatible steps or notes.", "orange"); return
            count = len(self.steps)
            y_offset = 0
            if self.steps:
                for i in range(len(self.steps)): self._calculate_node_size(i)
                max_y = max((s.get('y', 0) + s.get('_height', 60) / self.zoom_factor for s in self.steps), default=0)
                y_offset = max_y + 60
            for s in cleaned_steps:
                for key in ['on_success_goto_step', 'on_timeout_goto_step', 'on_count_reached_goto_step']:
                    if s.get(key): s[key] += count
                if s.get('on_success_action') == 'Next Step' and 'on_success_goto_step' in s: s['on_success_goto_step'] += count
                s['y'] = s.get('y', 50) + y_offset
            for n in notes: n['y'] = n.get('y', 50) + y_offset
            self.steps.extend(cleaned_steps); self.annotations.extend(notes); self.redraw_flowchart(); self.log(f"Appended {len(cleaned_steps)} steps and {len(notes)} notes from {os.path.basename(filepath)}.")
        except Exception as e: messagebox.showerror("Import Error", f"Failed to load or process file: {e}"); self.log(f"Import failed: {e}", "red")

    def reset_all(self):
        if messagebox.askyesno("Confirm Reset", "Are you sure you want to delete all steps and notes? This cannot be undone."): 
            self.destroy_all_overlays()
            self.steps.clear(); self.annotations.clear(); self.template_cache.clear()
            self.folder_image_cache.clear()
            self.selected_items = []; self.populate_properties_panel(); self.redraw_flowchart()
            self.log("Flowchart has been reset.", "orange")

    def copy_selection(self, event=None):
        if not self.selected_items or self.selected_items[0]['type'] != 'step': return
        self.clipboard = []
        indices = sorted([item['index'] for item in self.selected_items])
        min_x = min(self.steps[i].get('x', 0) for i in indices)
        min_y = min(self.steps[i].get('y', 0) for i in indices)
        self.clipboard_origin = (min_x, min_y)
        for index in indices: self.clipboard.append({"data": copy.deepcopy(self.steps[index]),"original_index": index})
        self.log(f"Copied {len(self.clipboard)} steps.")

    def paste_selection(self, event=None):
        if not self.clipboard: return

        # Determine paste location based on the last step, similar to add_step
        paste_anchor_x, paste_anchor_y = 50, 50 
        if self.steps:
            last_step = self.steps[-1]
            self._calculate_node_size(len(self.steps) - 1)
            paste_anchor_x = last_step.get('x', 50)
            paste_anchor_y = last_step.get('y', 50) + last_step.get('_height', 60) / self.zoom_factor + 40
        
        # Calculate the offset needed to move the copied block from its original position
        # to the new anchor position.
        x_offset = paste_anchor_x - self.clipboard_origin[0]
        y_offset = paste_anchor_y - self.clipboard_origin[1]

        new_steps, old_to_new_index_map, current_len = [], {}, len(self.steps)

        for i, clip_item in enumerate(self.clipboard):
            original_index = clip_item["original_index"]
            old_to_new_index_map[original_index] = current_len + i
            new_step = copy.deepcopy(clip_item["data"])
            
            # Apply the calculated offset to each pasted step
            new_step['x'] = new_step.get('x', 0) + x_offset
            new_step['y'] = new_step.get('y', 0) + y_offset
            
            new_steps.append(new_step)

        # Update flow control to maintain links within the pasted block
        for i, new_step in enumerate(new_steps):
            new_index = current_len + i
            original_index = next((k for k, v in old_to_new_index_map.items() if v == new_index), None)

            for action_key, goto_key in [('on_success_action', 'on_success_goto_step'), ('on_timeout_action', 'on_timeout_goto_step'), ('on_count_reached_action', 'on_count_reached_goto_step')]:
                if action_key not in new_step: continue
                
                action = new_step.get(action_key)
                if action == 'Go to Step':
                    target_original_index = new_step.get(goto_key, 1) - 1
                    if target_original_index in old_to_new_index_map:
                        new_step[goto_key] = old_to_new_index_map[target_original_index] + 1
                    else:
                        new_step[action_key] = 'Stop'
                elif action == 'Next Step':
                    if original_index is not None and (original_index + 1) in old_to_new_index_map:
                        # Successor step is also part of the pasted group, so link to it
                        new_step[action_key] = 'Go to Step'
                        new_step[goto_key] = old_to_new_index_map[original_index + 1] + 1
                    else:
                        # This step's successor is outside the pasted group, so default to Stop
                        new_step[action_key] = 'Stop'
        
        self.steps.extend(new_steps)
        self.log(f"Pasted {len(new_steps)} steps.")
        new_indices_range = range(len(self.steps) - len(new_steps), len(self.steps))
        self.selected_items = [{'type': 'step', 'index': i} for i in new_indices_range]
        self.populate_properties_panel(); self.redraw_flowchart()
        self.update_all_area_overlays()

    def delete_selected_from_key(self, event=None):
        if not self.selected_items: return
        num_items, item_type = len(self.selected_items), self.selected_items[0]['type']
        if not messagebox.askyesno("Confirm Delete", f"Are you sure you want to delete the {num_items} selected {item_type}s?"): return

        if item_type == 'step': self.delete_selected()
        elif item_type == 'note':
            indices_to_remove = sorted([item['index'] for item in self.selected_items], reverse=True)
            for index in indices_to_remove: self.annotations.pop(index)
            self.log(f"Removed {len(indices_to_remove)} notes.")
            self.selected_items = []
            self.populate_properties_panel(); self.redraw_flowchart()
    

