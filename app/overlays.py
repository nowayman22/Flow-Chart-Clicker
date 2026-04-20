import tkinter as tk

class OverlaysMixin:
    def toggle_step_show_area_flag(self):
        """Updates the 'show_area' flag for the currently selected step and refreshes overlays."""
        if not (self.selected_items and len(self.selected_items) == 1 and self.selected_items[0]['type'] == 'step'):
            return
        index = self.selected_items[0]['index']
        step = self.steps[index]
        
        if 'show_area' in self.properties_widgets:
            is_checked = self.properties_widgets['show_area'].get()
            step['show_area'] = is_checked
            self.update_all_area_overlays()

    def update_all_area_overlays(self):
        """Destroys all current overlays and redraws them based on global and per-step settings."""
        self.destroy_all_overlays()
        show_all = self.enable_all_show_area.get()
        
        for i, step in enumerate(self.steps):
            if step.get("area"):
                if show_all or step.get("show_area"):
                    self._draw_overlay_for_step(i)
    
    def _draw_overlay_for_step(self, index):
        """Creates and displays a single Toplevel window for a step's area."""
        if index in self.area_overlays:
            return
            
        step = self.steps[index]
        area = step.get("area")
        if not area:
            return

        try:
            x1, y1, x2, y2 = area
            width = x2 - x1
            height = y2 - y1

            if width <= 0 or height <= 0: return

            overlay = tk.Toplevel(self.root)
            overlay.overrideredirect(True)
            overlay.attributes("-topmost", True)
            overlay.attributes("-transparentcolor", "black")
            overlay.geometry(f"{width}x{height}+{x1}+{y1}")

            canvas = tk.Canvas(overlay, bg="black", highlightthickness=0)
            canvas.pack(fill=tk.BOTH, expand=True)

            border_color = self.current_theme.get('accent_teal', '#56B6C2')
            canvas.create_rectangle(0, 0, width, height, outline=border_color, width=3)
            
            self.area_overlays[index] = overlay
        except Exception as e:
            self.log(f"Error creating overlay for Step {index+1}: {e}", "red")

    def destroy_all_overlays(self):
        """Safely destroys all Toplevel windows used for area overlays."""
        for overlay in self.area_overlays.values():
            try:
                overlay.destroy()
            except tk.TclError:
                pass 
        self.area_overlays.clear()

