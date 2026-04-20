import tkinter as tk
from tkinter import ttk
import tkinter.font as tkfont
import math

class CanvasMixin:
    def redraw_flowchart(self):
        self.canvas.delete("all")
        self.canvas.config(bg=self.current_theme['canvas'])
        
        # --- NEW: Draw Grid ---
        if self.grid_visible.get():
            self._draw_grid()
        # --- END NEW ---

        for i in self.annotations: self.draw_annotation(i)
        if self.steps:
            for i in range(len(self.steps)): self._calculate_node_size(i)
            for i, step in enumerate(self.steps):
                self.draw_connection(i, 'on_success_action', 'on_success_goto_step')
                draw_timeout = False
                # --- FIX: Added 'Movement Detect' to ensure its timeout arrow is drawn ---
                if step.get('type') not in ['logical'] or step.get('logical_type') in ['Number', 'Wait', 'Movement Detect']:
                    draw_timeout = True
                if draw_timeout:
                    self.draw_connection(i, 'on_timeout_action', 'on_timeout_goto_step')
                if step.get('type') == 'logical' and step.get('logical_type') == 'Count':
                    self.draw_connection(i, 'on_count_reached_action', 'on_count_reached_goto_step')

            for i, step in enumerate(self.steps): self.draw_node(i)

        all_items_bbox = self.canvas.bbox("all")
        if all_items_bbox: self.canvas.config(scrollregion=(all_items_bbox[0]-50, all_items_bbox[1]-50, all_items_bbox[2]+50, all_items_bbox[3]+50))
        if self._marquee_data.get("rect"):
            rect = self._marquee_data["rect"]
            self.canvas.create_rectangle(rect, outline="#3399ff", width=2, dash=(4,2), tags="marquee_rect")

    def _draw_grid(self):
        """Draws the grid lines on the flowchart canvas based on current settings."""
        try:
            spacing = self.grid_spacing.get() * self.zoom_factor
            if spacing < 5: return # Avoid drawing too dense a grid

            canvas_width = self.canvas.winfo_width()
            canvas_height = self.canvas.winfo_height()
            
            scroll_region = self.canvas.bbox("all")
            if not scroll_region:
                x_start, y_start, x_end, y_end = 0, 0, canvas_width, canvas_height
            else:
                x_start, y_start, x_end, y_end = scroll_region

            x_start = min(x_start, self.canvas.canvasx(0))
            y_start = min(y_start, self.canvas.canvasy(0))
            x_end = max(x_end, self.canvas.canvasx(canvas_width))
            y_end = max(y_end, self.canvas.canvasy(canvas_height))

            # --- Calculate color based on opacity ---
            opacity = self.grid_opacity.get()
            grid_hex = self.current_theme.get('accent_grey', '#5C6370').lstrip('#')
            bg_hex = self.current_theme.get('canvas', '#21252B').lstrip('#')
            
            grid_rgb = tuple(int(grid_hex[i:i+2], 16) for i in (0, 2, 4))
            bg_rgb = tuple(int(bg_hex[i:i+2], 16) for i in (0, 2, 4))
            
            final_rgb = tuple(int(gc * opacity + bgc * (1 - opacity)) for gc, bgc in zip(grid_rgb, bg_rgb))
            final_color = self.rgb_to_hex(final_rgb)
            # --- End color calculation ---

            # Draw vertical lines
            for x in range(int(x_start // spacing * spacing), int(x_end), int(spacing)):
                self.canvas.create_line(x, y_start, x, y_end, fill=final_color, tags="grid_line")
            
            # Draw horizontal lines
            for y in range(int(y_start // spacing * spacing), int(y_end), int(spacing)):
                self.canvas.create_line(x_start, y, x_end, y, fill=final_color, tags="grid_line")
            
            self.canvas.tag_lower("grid_line")
        except (ValueError, tk.TclError):
            pass # Handles cases where spacing is invalid or window is not ready

    def _calculate_node_size(self, index):
        step = self.steps[index]; z = self.zoom_factor; title_font = tkfont.Font(family='Helvetica', size=int(9*z), weight='bold'); details_font = tkfont.Font(family='Helvetica', size=int(8*z))
        title = f"Step {index+1}: {step.get('name', 'Unnamed')}"
        action_text = step.get('action', '')
        step_type_display = step['type'].title()
        if step['type'] == 'location':
            step_type_display = 'Click / Press'
            action_text = step.get('action', 'Left Click')
            if action_text == 'Key Press':
                action_text = f"Press Key: {step.get('key_to_press', '')}"
        elif step['type'] == 'logical':
            if step.get('logical_type') == 'Number':
                action_text = step.get('expression', '> 0')
            elif step.get('logical_type') == 'Wait':
                action_text = f"Wait for {step.get('max_time', 0)}s"
            elif step.get('logical_type') == 'Type Text':
                if step.get('text_source') == 'GE Interface':
                    action_text = f"Type GE: {step.get('ge_data_field')}"
                else:
                    action_text = "Type Text"
            elif step.get('logical_type') == 'GE Name Inject':
                action_text = "GE Name Inject"
            else:
                action_text = step.get('logical_type', 'Execute')
        
        type_text = f"Type: {step_type_display}\nAction: {action_text}"; title_width = title_font.measure(title); details_width = max(details_font.measure(line) for line in type_text.split('\n'))
        padding_x, padding_y = 20 * z, 20 * z; node_width = max(title_width, details_width, 140*z) + padding_x; node_height = (title_font.metrics("linespace") + details_font.metrics("linespace") * 2) + padding_y
        step['_width'], step['_height'] = node_width, node_height

    def draw_node(self, index):
        step = self.steps[index]; x, y = step.get('x', 50), step.get('y', 50); z = self.zoom_factor; node_width, node_height = step.get('_width', 180*z), step.get('_height', 60*z)
        theme = self.current_theme
        node_colors = {"color": "#2c3a2c", "png": "#2a3a49", "location": "#4a4a4a", "logical": "#4d452c"}
        
        fill_color = node_colors.get(step['type'], "#555555")
        is_selected = any(item['type'] == 'step' and item['index'] == index for item in self.selected_items)
        border_color = "#2a9fd6" if is_selected else theme['node_border']; border_width = 3 if is_selected else 1; tag = f"step_{index}"
        self.canvas.create_rectangle(x*z, y*z, (x + node_width/z)*z, (y + node_height/z)*z, fill=fill_color, outline=border_color, width=border_width, tags=(tag, "node"))
        title = f"Step {index+1}: {step.get('name', 'Unnamed')}"; self.canvas.create_text((x + node_width/z/2)*z, (y + node_height/z * 0.25)*z, text=title, width=(node_width-15*z), justify=tk.CENTER, font=('Helvetica', int(9*z), 'bold'), tags=(tag, "node"), fill=theme['fg'])
        
        action_text = step.get('action', '')
        step_type_display = step['type'].title()
        if step['type'] == 'location':
            step_type_display = 'Click / Press'
            action_text = step.get('action', 'Left Click')
            if action_text == 'Key Press':
                action_text = f"Press Key: {step.get('key_to_press', '')}"
        elif step['type'] == 'logical':
            if step.get('logical_type') == 'Number':
                action_text = step.get('expression', '> 0')
            elif step.get('logical_type') == 'Wait':
                action_text = f"Wait for {step.get('max_time', 0)}s"
            elif step.get('logical_type') == 'Type Text':
                if step.get('text_source') == 'GE Interface':
                    action_text = f"Type GE: {step.get('ge_data_field')}"
                else:
                    action_text = f"Type: {step.get('text_to_type', '')[:15]}"
            elif step.get('logical_type') == 'GE Inject':
                if step.get('ge_inject_field') == 'Quantity':
                    action_text = f"Inject Qty: {step.get('ge_inject_quantity', '1')}"
                else: # Default to Name
                    action_text = f"Inject: {step.get('ge_inject_name', '')[:15]}"
            else:
                action_text = step.get('logical_type', 'Execute')

        type_text = f"Type: {step_type_display}\nAction: {action_text}"
        self.canvas.create_text((x + node_width/z/2)*z, (y + node_height/z * 0.65)*z, text=type_text, fill=theme['node_text_grey'], font=('Helvetica', int(8*z)), tags=(tag, "node"), justify=tk.CENTER)

        if self.running and index == self.current_step_index: self.canvas.create_oval((x-10)*z, (y-10)*z, (x+10)*z, (y+10)*z, fill=theme['status_green'], outline="", tags=("runtime_highlight", tag))

    def draw_annotation(self, note):
        x, y, w, h, z = note['x'], note['y'], note['width'], note['height'], self.zoom_factor; index = self.annotations.index(note)
        is_selected = any(item['type'] == 'note' and item['index'] == index for item in self.selected_items)
        border_color = "#2a9fd6" if is_selected else "#888888"; border_width = 3 if is_selected else 1; tag = f"note_{index}"
        opacity = note.get('opacity', '0% (Border Only)')
        fill_color = note['color'] if opacity != '0% (Border Only)' else ""
        stipple_map = {"100%": "", "75%": "gray75", "50%": "gray50", "25%": "gray25"}; stipple = stipple_map.get(opacity, "")
        self.canvas.create_rectangle(x*z, y*z, (x+w)*z, (y+h)*z, fill=fill_color, outline=border_color, width=border_width, stipple=stipple, tags=(tag, "annotation"))
        padding = 5 * z
        self.canvas.create_text((x*z + padding), (y*z + padding), text=note['text'], width=(w*z - padding*2),
                                font=('Helvetica', int(10*z), 'italic'), fill=self.current_theme['fg'],
                                justify=tk.LEFT, anchor=tk.NW, tags=(tag, "annotation"))
        if is_selected:
            handle_size = 8 * z; hx, hy = (x+w)*z - handle_size/2, (y+h)*z - handle_size/2
            self.canvas.create_rectangle(hx, hy, hx+handle_size, hy+handle_size, fill=border_color, outline='white', tags=(tag, "annotation", "resize_handle"))

    def _is_bidirectional(self, source_index, target_index):
        if not (0 <= target_index < len(self.steps)): return False
        target_step = self.steps[target_index]
        
        if target_step.get('on_success_action') == 'Go to Step' and target_step.get('on_success_goto_step') - 1 == source_index: return True
        if target_step.get('on_success_action') == 'Next Step' and target_index + 1 == source_index: return True

        is_timeout_step = target_step.get('type') not in ['logical'] or target_step.get('logical_type') in ['Number', 'Wait']
        if is_timeout_step and target_step.get('on_timeout_action') == 'Go to Step' and target_step.get('on_timeout_goto_step') - 1 == source_index: return True
        
        if target_step.get('type') == 'logical' and target_step.get('logical_type') == 'Count' and target_step.get('on_count_reached_action') == 'Go to Step' and target_step.get('on_count_reached_goto_step') - 1 == source_index: return True
        
        return False

    def draw_connection(self, source_index, action_key, goto_key):
        step = self.steps[source_index]; action = step.get(action_key); target_index = -1; line_style = {}
        if action == 'Next Step':
            if source_index + 1 < len(self.steps): target_index = source_index + 1; line_style = {'fill': '#5c7a96', 'dash': (10, 5), 'width': 1.5 * self.zoom_factor}
        elif action == 'Go to Step':
            target_index = step.get(goto_key, 1) - 1
            if 'success' in action_key: line_style = {'fill': self.current_theme['status_green'], 'width': 2.0 * self.zoom_factor}
            elif 'count_reached' in action_key: line_style = {'fill': self.current_theme['status_red'], 'dash': (8, 2, 2, 2), 'width': 2.0 * self.zoom_factor}
            elif 'fail' in action_key: line_style = {'fill': self.current_theme['status_red'], 'dash': (4, 4), 'width': 2.0 * self.zoom_factor}
            else: line_style = {'fill': self.current_theme['status_orange'], 'dash': (6, 4), 'width': 2.0 * self.zoom_factor}
        
        if 0 <= target_index < len(self.steps) and target_index != source_index:
            is_bidirectional = self._is_bidirectional(source_index, target_index)
            start_pos_center = self.get_node_center(source_index); target_center = self.get_node_center(target_index)
            target_step = self.steps[target_index]; z = self.zoom_factor; target_w = target_step.get('_width', 180*z)/z; target_h = target_step.get('_height', 60*z)/z
            start_pos_z = (start_pos_center[0] * z, start_pos_center[1] * z); target_center_z = (target_center[0] * z, target_center[1] * z)
            
            if is_bidirectional and source_index > target_index:
                dx, dy = target_center_z[0] - start_pos_z[0], target_center_z[1] - start_pos_z[1]; dist = math.hypot(dx, dy)
                if dist < 1: return
                mid_x, mid_y = start_pos_z[0] + dx*0.5, start_pos_z[1] + dy*0.5; perp_x, perp_y = -dy/dist, dx/dist
                curve_amount = dist * 0.2
                ctrl_point = (mid_x + curve_amount*perp_x, mid_y + curve_amount*perp_y)
                end_pos = self._get_line_to_node_edge(ctrl_point, target_center_z, target_w*z, target_h*z)
                self.canvas.create_line(start_pos_z, ctrl_point, end_pos, smooth=True, arrow=tk.LAST, **line_style)
            else:
                end_pos = self._get_line_to_node_edge(start_pos_z, target_center_z, target_w*z, target_h*z)
                self.canvas.create_line(start_pos_z, end_pos, arrow=tk.LAST, **line_style)

    def _get_line_to_node_edge(self, p1, p2, w, h):
        p1_x, p1_y = p1; p2_x, p2_y = p2; dx, dy = p2_x - p1_x, p2_y - p1_y
        if dx == 0 and dy == 0: return p2

        rect_ratio = h / w if w != 0 else float('inf')
        line_ratio = abs(dy / dx) if dx != 0 else float('inf')
        
        x, y = p2_x, p2_y

        if line_ratio < rect_ratio:
            if dx > 0: x, y = p2_x - w/2, p2_y - (w/2 * dy)/dx if dx!=0 else p2_y
            else: x, y = p2_x + w/2, p2_y + (w/2 * dy)/dx if dx!=0 else p2_y
        else:
            if dy > 0: x, y = p2_x - (h/2 * dx)/dy if dy!=0 else p2_x, p2_y - h/2
            else: x, y = p2_x + (h/2 * dx)/dy if dy!=0 else p2_x, p2_y + h/2
        
        return (x, y)

    def on_canvas_press(self, event):
        canvas_x, canvas_y = self.canvas.canvasx(event.x), self.canvas.canvasy(event.y)
        overlapping = self.canvas.find_overlapping(canvas_x, canvas_y, canvas_x, canvas_y)
        is_shift_pressed = (event.state & 0x0001) != 0

        clicked_item = None
        for item_id in reversed(overlapping):
            tags = self.canvas.gettags(item_id)
            if any(t.startswith("step_") or t.startswith("note_") for t in tags):
                for tag in tags:
                    if tag.startswith("step_") or tag.startswith("note_"):
                        item_type, index_str = tag.split('_')
                        index = int(index_str)
                        mode = "move"
                        if "resize_handle" in tags and item_type == 'note':
                            mode = "resize"
                        clicked_item = {'type': item_type, 'index': index, 'mode': mode}
                        break
            if clicked_item:
                break
        
        if clicked_item:
            self._drag_data["item"] = f"{clicked_item['type']}_{clicked_item['index']}"
            self._drag_data["start_x"] = canvas_x / self.zoom_factor
            self._drag_data["start_y"] = canvas_y / self.zoom_factor
            self._drag_data["mode"] = clicked_item['mode']

            is_already_selected = any(item['type'] == clicked_item['type'] and item['index'] == clicked_item['index'] for item in self.selected_items)
            
            if is_shift_pressed:
                if is_already_selected:
                    self.selected_items = [item for item in self.selected_items if not (item['type'] == clicked_item['type'] and item['index'] == clicked_item['index'])]
                else:
                    if not self.selected_items or self.selected_items[0]['type'] == clicked_item['type']:
                         self.selected_items.append(clicked_item)
                    else: 
                         self.selected_items = [clicked_item]
            else:
                if not is_already_selected:
                    self.selected_items = [clicked_item]
            
            # --- Store initial positions for dragging ---
            self._drag_data['initial_positions'] = []
            for item in self.selected_items:
                item_list = self.steps if item['type'] == 'step' else self.annotations
                item_obj = item_list[item['index']]
                pos_data = {'x': item_obj.get('x', 0), 'y': item_obj.get('y', 0)}
                if item['type'] == 'note':
                    pos_data['width'] = item_obj.get('width', 100)
                    pos_data['height'] = item_obj.get('height', 50)
                self._drag_data['initial_positions'].append(pos_data)

        else:
            if is_shift_pressed:
                self._marquee_data = {"x": canvas_x, "y": canvas_y, "rect": (canvas_x, canvas_y, canvas_x, canvas_y)}
            else:
                self.selected_items = []
                self._drag_data["item"] = None
        
        self.populate_properties_panel()
        self.redraw_flowchart()

    def on_drag_motion(self, event):
        if self._marquee_data.get("x") is not None:
            canvas_x, canvas_y = self.canvas.canvasx(event.x), self.canvas.canvasy(event.y)
            x1 = self._marquee_data["x"]
            y1 = self._marquee_data["y"]
            self._marquee_data["rect"] = (x1, y1, canvas_x, canvas_y)
            self.redraw_flowchart()
            return

        if self._drag_data["item"] is None or not self.selected_items:
            return

        canvas_x = self.canvas.canvasx(event.x) / self.zoom_factor
        canvas_y = self.canvas.canvasy(event.y) / self.zoom_factor
        
        # Calculate total delta from the start of the drag
        total_dx = canvas_x - self._drag_data["start_x"]
        total_dy = canvas_y - self._drag_data["start_y"]

        # Handle resize mode for a single selected note
        if self._drag_data['mode'] == 'resize' and len(self.selected_items) == 1 and self.selected_items[0]['type'] == 'note':
            note = self.annotations[self.selected_items[0]['index']]
            initial_dims = self._drag_data['initial_positions'][0]
            note['width'] = max(50, initial_dims['width'] + total_dx)
            note['height'] = max(30, initial_dims['height'] + total_dy)
        else:  # Move mode for one or more items
            for i, selected in enumerate(self.selected_items):
                item_list = self.steps if selected['type'] == 'step' else self.annotations
                item_obj = item_list[selected['index']]
                
                if i >= len(self._drag_data['initial_positions']): continue
                
                initial_pos = self._drag_data['initial_positions'][i]
                
                new_x = initial_pos['x'] + total_dx
                new_y = initial_pos['y'] + total_dy
                
                # Apply grid latching if enabled
                if self.grid_latching.get():
                    try:
                        spacing = self.grid_spacing.get()
                        if spacing > 0:
                            item_obj['x'] = round(new_x / spacing) * spacing
                            item_obj['y'] = round(new_y / spacing) * spacing
                        else:
                            item_obj['x'] = new_x; item_obj['y'] = new_y
                    except (ValueError, tk.TclError):
                        item_obj['x'] = new_x; item_obj['y'] = new_y
                else:
                    item_obj['x'] = new_x
                    item_obj['y'] = new_y
        
        self.redraw_flowchart()

    def on_drag_release(self, event):
        if self._marquee_data:
            x1, y1, x2, y2 = self._marquee_data["rect"]
            
            overlapping_ids = self.canvas.find_enclosed(min(x1, x2), min(y1, y2), max(x1, x2), max(y1, y2))
            
            newly_selected = []
            for item_id in overlapping_ids:
                tags = self.canvas.gettags(item_id)
                if any(t.startswith("step_") or t.startswith("note_") for t in tags):
                    for tag in tags:
                        if tag.startswith("step_") or tag.startswith("note_"):
                            item_type, index_str = tag.split('_')
                            item = {'type': item_type, 'index': int(index_str)}
                            if item not in self.selected_items and item not in newly_selected:
                                newly_selected.append(item)
                            break
            
            if newly_selected:
                first_new_type = newly_selected[0]['type']
                if not self.selected_items or self.selected_items[0]['type'] == first_new_type:
                    self.selected_items.extend([s for s in newly_selected if s['type'] == first_new_type])
                else: 
                    self.selected_items = [s for s in newly_selected if s['type'] == first_new_type]

            self._marquee_data = {}
            self.populate_properties_panel()
            self.redraw_flowchart()
            return

        self._drag_data = {"start_x": 0, "start_y": 0, "item": None, "mode": "move", "initial_positions": []}

    def get_node_center(self, index):
        step, z = self.steps[index], self.zoom_factor
        return (step.get('x', 50) + step.get('_width', 180*z)/z/2, step.get('y', 50) + step.get('_height', 60*z)/z/2)

    def setup_hotkeys(self):
        try: keyboard.add_hotkey('f2',lambda: self.start() if not self.running else self.stop()); keyboard.add_hotkey('f3',self.capture_from_hotkey); keyboard.add_hotkey('f4',self.select_area_mode)
        except Exception as e: self.log(f"Failed to register hotkeys: {e}", "red")
        self.root.bind("<Control-c>", self.copy_selection)
        self.root.bind("<Control-v>", self.paste_selection)
        self.root.bind("<Delete>", self.delete_selected_from_key)

    def _bind_mouse_scroll(self):
        def _on_mouse_wheel(event):
            is_ctrl, is_shift = (event.state & 0x0004) != 0, (event.state & 0x0001) != 0; scroll_dir = -1 if (event.num == 4 or event.delta > 0) else 1
            if is_ctrl: self.zoom_factor *= 1.1 if scroll_dir == -1 else 0.9; self.zoom_factor = max(0.2, min(3.0, self.zoom_factor)); self.redraw_flowchart()
            elif is_shift: self.canvas.xview_scroll(scroll_dir, "units")
            else: self.canvas.yview_scroll(scroll_dir, "units")
        self.canvas.bind("<MouseWheel>", _on_mouse_wheel); self.canvas.bind("<Button-4>", _on_mouse_wheel); self.canvas.bind("<Button-5>", _on_mouse_wheel)


