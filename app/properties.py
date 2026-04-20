import tkinter as tk
from tkinter import ttk, messagebox, filedialog, scrolledtext
from PIL import Image, ImageTk
import os
import copy

class PropertiesMixin:
    def populate_properties_panel(self):
        for widget in self.props_tab.winfo_children():
            widget.destroy()
        self.properties_widgets = {}

        if not self.selected_items:
            default_label = ttk.Label(self.props_tab, text="\n\nSelect a step or note in the flowchart\nto view and edit its properties.", justify=tk.CENTER, font=('Helvetica', 10))
            default_label.pack(expand=True, fill=tk.BOTH)
            if self.current_theme:
                default_label.config(foreground=self.current_theme['node_text_grey'])
            self.properties_widgets['default_label'] = default_label
            return
        
        container = ttk.Frame(self.props_tab)
        container.pack(fill="both", expand=True)
        self.properties_widgets['container'] = container
        
        if len(self.selected_items) == 1:
            item = self.selected_items[0]
            if item['type'] == 'step':
                self.populate_step_properties(container, item['index'])
            elif item['type'] == 'note':
                self.populate_note_properties(container, item['index'])
        elif len(self.selected_items) > 1 and self.selected_items[0]['type'] == 'step':
            self.populate_multi_select_properties(container)
        else: # Multiple notes selected
             self.populate_multi_select_properties(container)


        self.update_widget_colors_recursive(container, self.current_theme)
        
        if len(self.selected_items) == 1:
            item = self.selected_items[0]
            if item['type'] == 'step' and self.steps[item['index']]['type'] == 'color':
                if 'color_swatch' in self.properties_widgets:
                    self.properties_widgets['color_swatch'].config(bg=self.rgb_to_hex(self.steps[item['index']].get('rgb')))
            elif item['type'] == 'note':
                if 'note_swatch' in self.properties_widgets:
                    self.properties_widgets['note_swatch'].config(bg=self.annotations[item['index']].get('color'))

    def populate_multi_select_properties(self, container):
        num_selected = len(self.selected_items)
        item_type = self.selected_items[0]['type']
        
        tk.Label(container, text=f"{num_selected} {item_type.title()}s Selected", font=('Helvetica', 11, 'bold')).pack(pady=(0, 10), anchor='w')
        
        # --- Handle Notes ---
        if item_type == 'note':
            tk.Label(container, text="Multi-editing for notes is not yet available.\nOnly universal actions are available.", justify=tk.CENTER).pack(pady=5)
            action_btn_frm = tk.Frame(container)
            action_btn_frm.pack(fill=tk.X, pady=(15, 0))
            tk.Button(action_btn_frm, text="Delete Selected", font=('Helvetica', 9, 'bold'), command=self.delete_selected_from_key, relief=tk.FLAT).pack(side=tk.LEFT, expand=True, fill=tk.X, padx=2)
            # Spacer to push content up
            ttk.Frame(container).pack(fill=tk.BOTH, expand=True)
            return

        # --- Handle Steps ---
        selected_steps = [self.steps[item['index']] for item in self.selected_items]

        def get_common_value(prop_key, default_value):
            first_value = selected_steps[0].get(prop_key, default_value)
            if all(s.get(prop_key, default_value) == first_value for s in selected_steps):
                return first_value
            return self.MULTIPLE_VALUES

        # --- Universal Options Frame ---
        options_lf = ttk.LabelFrame(container, text="Options")
        options_lf.pack(fill=tk.X, pady=(5,10))
        options_lf.columnconfigure(1, weight=1)
        
        # Enable Logging Checkbox
        log_val = get_common_value('enable_logging', True)
        log_var = tk.BooleanVar()
        if log_val is not self.MULTIPLE_VALUES: log_var.set(log_val)
        self.properties_widgets['enable_logging'] = log_var
        log_cb = ttk.Checkbutton(options_lf, text="Enable Execution Log", variable=log_var)
        log_cb.grid(row=0, column=0, columnspan=2, sticky='w', pady=2, padx=5)
        if log_val is self.MULTIPLE_VALUES: log_cb.config(text="Enable Execution Log (mixed)")

        # Show Area Checkbox (if applicable)
        is_area_applicable = any(s['type'] in ['color', 'png'] or (s['type'] == 'logical' and s.get('logical_type') in ['Number', 'Movement Detect']) for s in selected_steps)
        if is_area_applicable:
            area_val = get_common_value('show_area', False)
            area_var = tk.BooleanVar()
            if area_val is not self.MULTIPLE_VALUES: area_var.set(area_val)
            self.properties_widgets['show_area'] = area_var
            area_cb = ttk.Checkbutton(options_lf, text="Show Area Overlay", variable=area_var, command=self.update_all_area_overlays)
            area_cb.grid(row=1, column=0, columnspan=2, sticky='w', pady=2, padx=5)
            if area_val is self.MULTIPLE_VALUES: area_cb.config(text="Show Area Overlay (mixed)")

        # --- Type-Specific Flow Control Frame ---
        first_step_type = selected_steps[0]['type']
        if all(s['type'] == first_step_type for s in selected_steps):
            flow_lf = ttk.LabelFrame(container, text=f"{first_step_type.title()} Flow Control")
            flow_lf.pack(fill=tk.X)
            flow_lf.columnconfigure(1, weight=1)
            row = 0

            # Collect properties based on type
            props_to_show = ['delay_after', 'on_success_action']
            if first_step_type in ['color', 'png']:
                props_to_show.extend(['timeout', 'on_timeout_action'])
            
            for prop in props_to_show:
                default_map = {'delay_after': 1.0, 'on_success_action': 'Next Step', 'timeout': 5.0, 'on_timeout_action': 'Stop'}
                value = get_common_value(prop, default_map.get(prop))
                label_text = prop.replace('_', ' ').title()
                if "On Success" in label_text: label_text = "On Success:"
                if "On Timeout" in label_text: label_text = "On Timeout:"
                ttk.Label(flow_lf, text=f"{label_text}").grid(row=row, column=0, sticky='w', pady=3, padx=5)
                
                if prop.endswith('_action'):
                    options = ["Next Step", "Go to Step", "Stop"]
                    var = tk.StringVar(value=value if value != self.MULTIPLE_VALUES else "")
                    self.properties_widgets[prop] = var
                    ttk.OptionMenu(flow_lf, var, "" if value == self.MULTIPLE_VALUES else value, *options).grid(row=row, column=1, sticky='ew', padx=5)
                else: # Entry-based
                    var = tk.StringVar(value=value)
                    self.properties_widgets[prop] = var
                    entry = ttk.Entry(flow_lf, textvariable=var)
                    if value == self.MULTIPLE_VALUES: entry.config(foreground=self.current_theme['node_text_grey'])
                    entry.grid(row=row, column=1, sticky='ew', padx=5)
                row += 1
        else:
             tk.Label(container, text="Selected steps are of different types.\nOnly universal options are available.", justify=tk.CENTER).pack(pady=5)
        
        # --- Action Buttons ---
        tk.Button(container, text="Apply Changes to Selected", font=('Helvetica', 9, 'bold'), command=self.apply_multi_properties_changes, relief=tk.FLAT).pack(fill=tk.X, pady=(15, 5), ipady=4)
        
        action_btn_frm = tk.Frame(container)
        action_btn_frm.pack(fill=tk.X, pady=(0, 0))
        btn_pack_style = {'side': tk.LEFT, 'expand': True, 'fill': tk.X, 'padx': 2}
        tk.Button(action_btn_frm, text="Duplicate Selected", font=('Helvetica', 9, 'bold'), command=self.duplicate_selected, relief=tk.FLAT).pack(**btn_pack_style)
        tk.Button(action_btn_frm, text="Delete Selected", font=('Helvetica', 9, 'bold'), command=self.delete_selected_from_key, relief=tk.FLAT).pack(**btn_pack_style)

        # --- Spacer to push all content to the top ---
        ttk.Frame(container).pack(fill=tk.BOTH, expand=True)

    def populate_step_properties(self, container, i):
        step = self.steps[i]
        tk.Label(container, text=f"Properties for Step {i+1}", font=('Helvetica', 11, 'bold')).grid(row=0, columnspan=3, sticky='w', pady=(0,10))
        tk.Label(container, text="Name:").grid(row=1, column=0, sticky='w', pady=2); w = tk.Entry(container); w.insert(0, step.get('name', '')); w.grid(row=1, column=1, columnspan=2, sticky='ew', pady=2); self.properties_widgets['name_entry'] = w
        
        action_row = 2
        tk.Label(container, text="Action:").grid(row=action_row, column=0, sticky='w', pady=2)
        action_frm = tk.Frame(container); action_frm.grid(row=action_row, column=1, columnspan=2, sticky='w')

        details_label_text = {'location': 'Click / Press Details'}
        details_label = details_label_text.get(step['type'], f"{step['type'].title()} Details")
        
        details_lf = tk.LabelFrame(container, text=details_label, padx=5, pady=5); details_lf.grid(row=3, columnspan=3, sticky='ew', pady=10)
        self.properties_widgets['details_lf'] = details_lf

        options_lf = tk.LabelFrame(container, text="Options", padx=5, pady=5); options_lf.grid(row=5, columnspan=3, sticky='ew', pady=(0, 10))
        w = tk.BooleanVar(value=step.get('enable_logging', True)); self.properties_widgets['enable_logging'] = w
        ttk.Checkbutton(options_lf, text="Enable Execution Log", variable=w).pack(side=tk.LEFT, anchor='w')

        uses_area = step['type'] in ['color', 'png'] or (step['type'] == 'logical' and step.get('logical_type') in ['Number', 'Movement Detect'])
        if uses_area:
            w = tk.BooleanVar(value=step.get('show_area', False)); self.properties_widgets['show_area'] = w
            ttk.Checkbutton(options_lf, text="Show Area", variable=w, command=self.toggle_step_show_area_flag).pack(side=tk.LEFT, anchor='w', padx=(10, 0))

        flow_lf = tk.LabelFrame(container, text="Flow Control", padx=5, pady=5); flow_lf.grid(row=6, columnspan=3, sticky='ew', pady=0)

        if step['type'] == 'logical':
            logical_type_var = tk.StringVar(value=step.get('logical_type', 'Count')); self.properties_widgets['logical_type'] = logical_type_var
            def _update_logical_details_frame(flow_control_frame):
                for widget in details_lf.winfo_children(): widget.destroy()
                for w in flow_control_frame.grid_slaves():
                    if int(w.grid_info()["row"]) > 1: w.destroy()
                timeout_widgets = self.properties_widgets.get('timeout_widgets', [])
                selected_type = logical_type_var.get(); step['logical_type'] = selected_type
                if selected_type in ['Number', 'Movement Detect']: [w.grid() for w in timeout_widgets]
                else: [w.grid_remove() for w in timeout_widgets]
                if selected_type == 'Count':
                    tk.Label(details_lf, text="Current Count:").grid(row=0, column=0, sticky='w', pady=5); w = tk.Label(details_lf, text=str(step.get('counter_value', 0)), font=('Consolas', 10, 'bold')); w.grid(row=0, column=1, sticky='w', padx=5); self.properties_widgets['counter_display'] = w; w = tk.Button(details_lf, text="Reset Count", font=('Helvetica', 9), command=self.reset_logical_counter, relief=tk.FLAT); w.grid(row=0, column=2, padx=10)
                    tk.Label(details_lf, text="Max Count (0=inf):").grid(row=1, column=0, sticky='w', pady=5); w = tk.Entry(details_lf, width=8); w.insert(0, str(step.get('max_count', 0))); w.grid(row=1, column=1, sticky='w', padx=5); self.properties_widgets['max_count'] = w
                    w = tk.BooleanVar(value=step.get('reset_on_start', False)); self.properties_widgets['reset_on_start'] = w; ttk.Checkbutton(details_lf, text="Reset on Start (F2)", variable=w).grid(row=2, column=0, columnspan=3, sticky='w', pady=5)
                    w = tk.BooleanVar(value=step.get('reset_on_reach', False)); self.properties_widgets['reset_on_reach'] = w; ttk.Checkbutton(details_lf, text="Reset on Count Reached", variable=w).grid(row=3, column=0, columnspan=3, sticky='w', pady=5)
                    
                    tk.Label(flow_control_frame, text="On Count Reached:").grid(row=2,column=0,sticky='w',pady=2); w = tk.StringVar(value=step.get('on_count_reached_action')); self.properties_widgets['on_count_reached_action'] = w; tk.OptionMenu(flow_control_frame,w,"Next Step","Go to Step", "Stop").grid(row=2,column=1,sticky='ew'); w = tk.Entry(flow_control_frame,width=5); w.insert(0,str(step.get('on_count_reached_goto_step'))); w.grid(row=2,column=2,padx=5); self.properties_widgets['on_count_reached_goto_step'] = w
                    tk.Label(flow_control_frame, text="Delay (s):").grid(row=2, column=3, padx=(10,0)); w = tk.Entry(flow_control_frame, width=7); w.insert(0, str(step.get('on_count_reached_delay'))); w.grid(row=2, column=4); self.properties_widgets['on_count_reached_delay'] = w
                elif selected_type == 'Wait':
                    tk.Label(details_lf, text="Wait Duration (s):").grid(row=0, column=0, sticky='w', pady=5); w = tk.Entry(details_lf, width=8); w.insert(0, str(step.get('max_time', 0))); w.grid(row=0, column=1, sticky='w', padx=5); self.properties_widgets['max_time'] = w
                    w = tk.Button(details_lf, text="Reset Wait Timer", font=('Helvetica', 9), command=self.reset_timer, relief=tk.FLAT); w.grid(row=0, column=2, padx=10)
                    w = tk.BooleanVar(value=step.get('reset_on_start', False)); self.properties_widgets['reset_on_start'] = w; ttk.Checkbutton(details_lf, text="Reset on Start (F2)", variable=w).grid(row=1, column=0, columnspan=3, sticky='w', pady=5)
                elif selected_type == 'Type Text':
                    source_frame = ttk.Frame(details_lf); source_frame.grid(row=0, column=0, columnspan=3, sticky='w', pady=2)
                    source_var = tk.StringVar(value=step.get('text_source', 'Static Text')); self.properties_widgets['text_source'] = source_var
                    
                    static_entry = ttk.Entry(details_lf, width=25); self.properties_widgets['text_to_type'] = static_entry; static_entry.insert(0, str(step.get('text_to_type', '')))
                    ge_data_options = ["Item Name", "Quantity", "Calculated Buy Price", "Calculated Sell Price", "Calculated Buy Total", "Calculated Sell Total"]
                    ge_data_var = tk.StringVar(value=step.get('ge_data_field', 'Calculated Buy Price')); self.properties_widgets['ge_data_field'] = ge_data_var
                    ge_data_combo = ttk.Combobox(details_lf, textvariable=ge_data_var, values=ge_data_options, state='readonly', width=22)
                    
                    def _toggle_source_widgets(*args):
                        if source_var.get() == 'GE Interface':
                            static_entry.grid_remove(); ge_data_combo.grid(row=1, column=1, sticky='ew', padx=5)
                        else:
                            ge_data_combo.grid_remove(); static_entry.grid(row=1, column=1, sticky='ew', padx=5)
                    
                    ttk.Radiobutton(source_frame, text="Static Text:", variable=source_var, value="Static Text", command=_toggle_source_widgets).pack(side=tk.LEFT)
                    ttk.Radiobutton(source_frame, text="From GE Interface:", variable=source_var, value="GE Interface", command=_toggle_source_widgets).pack(side=tk.LEFT, padx=10)
                    
                    enter_frame = ttk.Frame(details_lf)
                    enter_frame.grid(row=2, column=1, sticky='ew', pady=5)
                    
                    press_enter_var = tk.BooleanVar(value=step.get('press_enter', False))
                    self.properties_widgets['press_enter'] = press_enter_var
                    
                    delay_label = ttk.Label(enter_frame, text="Delay (s):")
                    delay_entry = ttk.Entry(enter_frame, width=7)
                    delay_entry.insert(0, str(step.get('enter_press_delay', 0.1)))
                    self.properties_widgets['enter_press_delay'] = delay_entry
                    
                    def _toggle_enter_delay(*args):
                        if press_enter_var.get():
                            delay_label.pack(side=tk.LEFT, padx=(10, 2))
                            delay_entry.pack(side=tk.LEFT)
                        else:
                            delay_label.pack_forget()
                            delay_entry.pack_forget()

                    press_enter_cb = ttk.Checkbutton(enter_frame, text="Press Enter after", variable=press_enter_var, command=_toggle_enter_delay)
                    press_enter_cb.pack(side=tk.LEFT)
                    
                    press_enter_var.trace_add('write', _toggle_enter_delay)
                    _toggle_enter_delay() 
                    
                    details_lf.columnconfigure(1, weight=1)
                    _toggle_source_widgets()
                elif selected_type == 'GE Inject':
                    field_var = tk.StringVar(value=step.get('ge_inject_field', 'Name'))
                    self.properties_widgets['ge_inject_field'] = field_var

                    field_frame = ttk.Frame(details_lf)
                    field_frame.grid(row=0, column=0, columnspan=2, sticky='w', pady=(0, 5))
                    
                    name_label = ttk.Label(details_lf, text="Item Name:")
                    name_entry = ttk.Entry(details_lf, width=25)
                    name_entry.insert(0, str(step.get('ge_inject_name', '')))
                    self.properties_widgets['ge_inject_name'] = name_entry
                    
                    qty_label = ttk.Label(details_lf, text="Quantity:")
                    qty_entry = ttk.Entry(details_lf, width=25)
                    qty_entry.insert(0, str(step.get('ge_inject_quantity', '1')))
                    self.properties_widgets['ge_inject_quantity'] = qty_entry

                    def _toggle_inject_widgets(*args):
                        if field_var.get() == 'Quantity':
                            name_label.grid_remove(); name_entry.grid_remove()
                            qty_label.grid(row=1, column=0, sticky='w', pady=5)
                            qty_entry.grid(row=1, column=1, sticky='ew', padx=5)
                        else: # Name
                            qty_label.grid_remove(); qty_entry.grid_remove()
                            name_label.grid(row=1, column=0, sticky='w', pady=5)
                            name_entry.grid(row=1, column=1, sticky='ew', padx=5)
                    
                    ttk.Radiobutton(field_frame, text="Name", variable=field_var, value="Name", command=_toggle_inject_widgets).pack(side=tk.LEFT)
                    ttk.Radiobutton(field_frame, text="Quantity", variable=field_var, value="Quantity", command=_toggle_inject_widgets).pack(side=tk.LEFT, padx=10)

                    refresh_var = tk.BooleanVar(value=step.get('ge_inject_refresh', False))
                    self.properties_widgets['ge_inject_refresh'] = refresh_var
                    ttk.Checkbutton(details_lf, text="Refresh Price After Inject", variable=refresh_var).grid(row=2, column=0, columnspan=2, sticky='w', pady=(5, 0))

                    details_lf.columnconfigure(1, weight=1)
                    _toggle_inject_widgets()
                elif selected_type == 'Settings Inject':
                    setting_map = {
                        "Location Offset (±px)": 'loc_offset_variance', "Speed Variance (±s)": 'speed_variance',
                        "Hold Variance (±s)": 'hold_duration_variance', "Scan Interval (s)": 'scan_interval',
                        "Base Hold Duration (s)": 'hold_duration'
                    }
                    setting_var = tk.StringVar(value=step.get('inject_setting_name', next(iter(setting_map))))
                    self.properties_widgets['inject_setting_name'] = setting_var
                    
                    tk.Label(details_lf, text="Setting to Change:").grid(row=0, column=0, sticky='w', pady=5)
                    ttk.OptionMenu(details_lf, setting_var, setting_var.get(), *setting_map.keys()).grid(row=0, column=1, sticky='ew', padx=5)
                    
                    tk.Label(details_lf, text="New Value:").grid(row=1, column=0, sticky='w', pady=5)
                    w = tk.Entry(details_lf, width=15); w.insert(0, str(step.get('inject_setting_value', ''))); w.grid(row=1, column=1, sticky='ew', padx=5)
                    self.properties_widgets['inject_setting_value'] = w
                    details_lf.columnconfigure(1, weight=1)
                elif selected_type == 'Number':
                    tk.Label(details_lf, text="Expression:").grid(row=0, column=0, sticky='w'); w=tk.Entry(details_lf, width=15); w.insert(0, str(step.get('expression', ''))); w.grid(row=0, column=1, columnspan=2, sticky='ew'); self.properties_widgets['expression'] = w; tk.Label(details_lf, text="e.g., > 100").grid(row=0, column=3, sticky='w', padx=5)
                    tk.Label(details_lf, text="Image Mode:").grid(row=1, column=0, sticky='w', pady=5); w = tk.StringVar(value=step.get('image_mode', 'Grayscale')); self.properties_widgets['number_image_mode'] = w; tk.OptionMenu(details_lf, w, 'Grayscale', 'Color', 'Binary (B&W)').grid(row=1, column=1, columnspan=3, sticky='ew')
                    tk.Label(details_lf, text="OEM:").grid(row=2, column=0, sticky='w', pady=2); oem_var = tk.StringVar(value=step.get('oem_mode', "3: Default, based on what is available.")); self.properties_widgets['oem_mode'] = oem_var; w = ttk.OptionMenu(details_lf, oem_var, oem_var.get(), *self.oem_options.keys()); w.grid(row=2, column=1, columnspan=3, sticky='ew')
                    tk.Label(details_lf, text="PSM:").grid(row=3, column=0, sticky='w', pady=2); psm_var = tk.StringVar(value=step.get('psm_mode', "6: Assume a single uniform block of text.")); self.properties_widgets['psm_mode'] = psm_var; w = ttk.OptionMenu(details_lf, psm_var, psm_var.get(), *self.psm_options.keys()); w.grid(row=3, column=1, columnspan=3, sticky='ew')
                    area_btn_frame = tk.Frame(details_lf); area_btn_frame.grid(row=4, column=0, columnspan=4, sticky='w', pady=(5,0))
                    area_text = f"Area: {step['area'][2]-step['area'][0]}x{step['area'][3]-step['area'][1]}" if step.get('area') else "Area: Global"
                    w = tk.Button(area_btn_frame, text=area_text, command=self.select_area_for_step, font=('Helvetica', 9), relief=tk.FLAT); w.pack(side=tk.LEFT, padx=(0, 2)); self.properties_widgets['area_btn'] = w
                    tk.Button(area_btn_frame, text="Full Screen", command=self.set_step_area_to_fullscreen, font=('Helvetica', 9), relief=tk.FLAT).pack(side=tk.LEFT, padx=(0, 2))
                    tk.Button(area_btn_frame, text="Use Global", command=self.set_step_area_to_global, font=('Helvetica', 9), relief=tk.FLAT).pack(side=tk.LEFT)
                    details_lf.columnconfigure(1, weight=1)
                elif selected_type == 'Movement Detect':
                    tk.Label(details_lf, text="Stillness Tolerance (%):").grid(row=0, column=0, sticky='w', pady=5)
                    w = tk.Entry(details_lf, width=8)
                    w.insert(0, str(step.get('movement_tolerance', 5.0)))
                    w.grid(row=0, column=1, sticky='w', padx=5)
                    self.properties_widgets['movement_tolerance'] = w
                    
                    w = tk.BooleanVar(value=step.get('reset_on_start', True))
                    self.properties_widgets['reset_on_start'] = w
                    ttk.Checkbutton(details_lf, text="Reset Comparison on Start (F2)", variable=w).grid(row=1, column=0, columnspan=3, sticky='w', pady=5)

                    area_btn_frame = tk.Frame(details_lf)
                    area_btn_frame.grid(row=2, column=0, columnspan=4, sticky='w', pady=(5,0))
                    area_text = f"Area: {step['area'][2]-step['area'][0]}x{step['area'][3]-step['area'][1]}" if step.get('area') else "Area: Global"
                    w = tk.Button(area_btn_frame, text=area_text, command=self.select_area_for_step, font=('Helvetica', 9), relief=tk.FLAT)
                    w.pack(side=tk.LEFT, padx=(0, 2))
                    self.properties_widgets['area_btn'] = w
                    tk.Button(area_btn_frame, text="Full Screen", command=self.set_step_area_to_fullscreen, font=('Helvetica', 9), relief=tk.FLAT).pack(side=tk.LEFT, padx=(0, 2))
                    tk.Button(area_btn_frame, text="Use Global", command=self.set_step_area_to_global, font=('Helvetica', 9), relief=tk.FLAT).pack(side=tk.LEFT)
                    details_lf.columnconfigure(1, weight=1)
                self.update_widget_colors_recursive(container, self.current_theme)
            
            command = lambda: _update_logical_details_frame(flow_lf)
            logical_radios_frm_1 = tk.Frame(action_frm); logical_radios_frm_1.pack(fill=tk.X, anchor='w')
            logical_radios_frm_2 = tk.Frame(action_frm); logical_radios_frm_2.pack(fill=tk.X, anchor='w')
            tk.Radiobutton(logical_radios_frm_1, text="Count", variable=logical_type_var, value="Count", command=command).pack(side=tk.LEFT, padx=(0,5))
            tk.Radiobutton(logical_radios_frm_1, text="Wait/Delay", variable=logical_type_var, value="Wait", command=command).pack(side=tk.LEFT, padx=5)
            tk.Radiobutton(logical_radios_frm_1, text="Type Text", variable=logical_type_var, value="Type Text", command=command).pack(side=tk.LEFT, padx=5)
            tk.Radiobutton(logical_radios_frm_2, text="GE Inject", variable=logical_type_var, value="GE Inject", command=command).pack(side=tk.LEFT, padx=(0,5))
            tk.Radiobutton(logical_radios_frm_2, text="Settings Inject", variable=logical_type_var, value="Settings Inject", command=command).pack(side=tk.LEFT, padx=5)
            if PYTESSERACT_AVAILABLE: tk.Radiobutton(logical_radios_frm_2, text="Number", variable=logical_type_var, value="Number", command=command).pack(side=tk.LEFT, padx=5)
            tk.Radiobutton(logical_radios_frm_2, text="Movement", variable=logical_type_var, value="Movement Detect", command=command).pack(side=tk.LEFT, padx=5)
        else:
            action_var = tk.StringVar(value=step.get('action')); self.properties_widgets['action'] = action_var
            if step['type'] == 'location':
                def _update_location_details():
                    is_key_press = action_var.get() == 'Key Press'
                    is_click_only = action_var.get() == 'Click Only'
                    for widget_key in ['key_press_entry', 'key_press_label']:
                        if self.properties_widgets.get(widget_key):
                            self.properties_widgets[widget_key].grid() if is_key_press else self.properties_widgets[widget_key].grid_remove()
                    for widget_key in ['get_loc_btn', 'coords_label']:
                        if self.properties_widgets.get(widget_key):
                             self.properties_widgets[widget_key].grid_remove() if is_key_press or is_click_only else self.properties_widgets[widget_key].grid()
                
                tk.Radiobutton(action_frm, text="Left Click", variable=action_var, value='Left Click', command=_update_location_details).pack(side=tk.LEFT)
                tk.Radiobutton(action_frm, text="Right Click", variable=action_var, value='Right Click', command=_update_location_details).pack(side=tk.LEFT, padx=5)
                tk.Radiobutton(action_frm, text="Move Only", variable=action_var, value='Move Only', command=_update_location_details).pack(side=tk.LEFT, padx=5)
                tk.Radiobutton(action_frm, text="Click Only", variable=action_var, value='Click Only', command=_update_location_details).pack(side=tk.LEFT, padx=5)
                tk.Radiobutton(action_frm, text="Key Press", variable=action_var, value='Key Press', command=_update_location_details).pack(side=tk.LEFT, padx=5)
                self.properties_widgets['get_loc_btn'], self.properties_widgets['coords_label'] = None, None

            elif step['type'] == 'png' or step['type'] == 'color':
                def _update_details_for_action():
                    action = action_var.get()
                    is_count_mode = (step['type'] == 'png' and action == 'PNG Count') or \
                                    (step['type'] == 'color' and action == 'Color Count')
                    
                    if 'count_details_lf' in self.properties_widgets:
                        self.properties_widgets['count_details_lf'].grid() if is_count_mode else self.properties_widgets['count_details_lf'].grid_remove()
                    
                    if 'timeout_action_label' in self.properties_widgets:
                        self.properties_widgets['timeout_action_label'].config(text="On Fail:") if is_count_mode else self.properties_widgets['timeout_action_label'].config(text="On Timeout:")
                    
                    if 'timeout_widgets' in self.properties_widgets:
                        timeout_label, w_timeout = self.properties_widgets['timeout_widgets'][3], self.properties_widgets['timeout_widgets'][4]
                        if is_count_mode:
                            timeout_label.grid_remove()
                            w_timeout.grid_remove()
                        else:
                            timeout_label.grid()
                            w_timeout.grid()

                tk.Radiobutton(action_frm, text="Left Click", variable=action_var, value='Click Object', command=_update_details_for_action).pack(side=tk.LEFT)
                tk.Radiobutton(action_frm, text="Right Click", variable=action_var, value='Right Click', command=_update_details_for_action).pack(side=tk.LEFT, padx=5)
                tk.Radiobutton(action_frm, text="Detect Only", variable=action_var, value='Detect Object', command=_update_details_for_action).pack(side=tk.LEFT, padx=5)
                if step['type'] == 'png':
                    tk.Radiobutton(action_frm, text="PNG Count", variable=action_var, value='PNG Count', command=_update_details_for_action).pack(side=tk.LEFT, padx=5)
                if step['type'] == 'color':
                    tk.Radiobutton(action_frm, text="Color Count", variable=action_var, value='Color Count', command=_update_details_for_action).pack(side=tk.LEFT, padx=5)

        if step['type'] == 'color':
            pixel_detect_var = tk.BooleanVar(value=step.get('pixel_detect_enabled', False))
            self.properties_widgets['pixel_detect_enabled'] = pixel_detect_var
            
            picker_frame = tk.Frame(details_lf)
            picker_frame.grid(row=0, column=0, columnspan=4, sticky='w')
            tk.Button(picker_frame,text="Pick (F3)", font=('Helvetica', 9), command=lambda: self.enter_f3_mode('pick_color'), relief=tk.FLAT).pack(side=tk.LEFT)
            w = tk.Label(picker_frame,text=" ",bg=self.rgb_to_hex(step.get('rgb')),relief=tk.SUNKEN,width=3); w.pack(side=tk.LEFT, padx=4); self.properties_widgets['color_swatch'] = w
            
            coords_text = f"Pixel: {step.get('pixel_coords')}" if step.get('pixel_coords') else "Pixel: Not Set"
            pixel_coords_label = tk.Label(picker_frame, text=coords_text); self.properties_widgets['pixel_coords_label'] = pixel_coords_label

            tk.Label(details_lf,text="Tolerance:").grid(row=1,column=0, sticky='w', pady=2); w=tk.Entry(details_lf,width=7); w.insert(0,str(step.get('tolerance'))); w.grid(row=1,column=1, sticky='w', padx=2); self.properties_widgets['tolerance'] = w
            tk.Label(details_lf, text="Min Area (px):").grid(row=1, column=2, sticky='w', padx=(10,2)); w=tk.Entry(details_lf, width=7); w.insert(0, str(step.get('min_pixel_area', 10))); w.grid(row=1, column=3, sticky='w', padx=2); self.properties_widgets['min_pixel_area'] = w
            
            tk.Label(details_lf, text="Color Space:").grid(row=2, column=0, sticky='w', pady=5); w = tk.StringVar(value=step.get('color_space', 'HSV')); self.properties_widgets['color_space'] = w; 
            tk.OptionMenu(details_lf, w, 'HSV', 'RGB').grid(row=2, column=1, columnspan=3, sticky='ew')
            
            area_btn_frame = tk.Frame(details_lf); area_btn_frame.grid(row=3, column=0, columnspan=4, sticky='w', pady=(5,0))
            area_text = f"Area: {step['area'][2]-step['area'][0]}x{step['area'][3]-step['area'][1]}" if step.get('area') else "Area: Global"
            w_area_btn = tk.Button(area_btn_frame, text=area_text, command=self.select_area_for_step, font=('Helvetica', 9), relief=tk.FLAT); w_area_btn.pack(side=tk.LEFT, padx=(0, 2)); self.properties_widgets['area_btn'] = w_area_btn
            tk.Button(area_btn_frame, text="Full Screen", command=self.set_step_area_to_fullscreen, font=('Helvetica', 9), relief=tk.FLAT).pack(side=tk.LEFT, padx=(0, 2))
            tk.Button(area_btn_frame, text="Use Global", command=self.set_step_area_to_global, font=('Helvetica', 9), relief=tk.FLAT).pack(side=tk.LEFT)

            def _update_color_details_ui(*args):
                is_pixel_mode = pixel_detect_var.get()
                if is_pixel_mode:
                    area_btn_frame.grid_remove()
                    pixel_coords_label.pack(side=tk.LEFT, padx=10)
                else:
                    area_btn_frame.grid()
                    pixel_coords_label.pack_forget()
                
                for widget in details_lf.grid_slaves():
                    if widget.grid_info()['row'] == 1 and widget.grid_info()['column'] in [2, 3]:
                        widget.grid_remove() if is_pixel_mode else widget.grid()


            ttk.Checkbutton(details_lf, text="Pixel Detect Mode", variable=pixel_detect_var, command=_update_color_details_ui).grid(row=4, column=0, columnspan=4, sticky='w', pady=(10,0))
            _update_color_details_ui()

        elif step['type'] == 'png':
            png_mode_frm = tk.Frame(details_lf); png_mode_frm.grid(row=0, columnspan=4, sticky='w')
            w = tk.StringVar(value=step.get('mode')); self.properties_widgets['png_mode'] = w
            tk.Radiobutton(png_mode_frm,text="File",variable=w,value='file').pack(side=tk.LEFT)
            tk.Radiobutton(png_mode_frm,text="Folder",variable=w,value='folder').pack(side=tk.LEFT, padx=10)
            
            path_frm = tk.Frame(details_lf); path_frm.grid(row=1, columnspan=4, sticky='ew', pady=5)
            tk.Button(path_frm,text="Snip",command=self.snip_image_for_step, font=('Helvetica', 9), relief=tk.FLAT).pack(side=tk.LEFT, padx=(0,5))
            tk.Button(path_frm,text="Browse",command=self.browse_for_step, font=('Helvetica', 9), relief=tk.FLAT).pack(side=tk.LEFT)
            w = tk.Label(path_frm,text=os.path.basename(step.get('path'))or"No path set",anchor='w',wraplength=180,justify='left'); w.pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
            self.properties_widgets['path'] = w
            
            w_preview = tk.Label(details_lf, text="No Preview Available")
            w_preview.grid(row=2, column=0, columnspan=4, sticky='ew', pady=5)
            self.properties_widgets['png_preview'] = w_preview
            self._update_png_preview(step)

            tk.Label(details_lf,text="Threshold:").grid(row=3,column=0,sticky='w', pady=2)
            w=tk.Entry(details_lf,width=7); w.insert(0,str(step.get('threshold'))); w.grid(row=3,column=1)
            self.properties_widgets['threshold'] = w
            
            tk.Label(details_lf, text="Image Mode:").grid(row=4, column=0, sticky='w', pady=5)
            w = tk.StringVar(value=step.get('image_mode', 'Grayscale')); self.properties_widgets['image_mode'] = w
            tk.OptionMenu(details_lf, w, 'Grayscale', 'Color', 'Binary (B&W)').grid(row=4, column=1, columnspan=2, sticky='ew')

            w = tk.BooleanVar(value=step.get('find_first_match', False)); self.properties_widgets['find_first_match'] = w
            ttk.Checkbutton(details_lf, text="Fast Mode (Find First Match)", variable=w).grid(row=5, column=0, columnspan=3, sticky='w', pady=2)
            
            area_btn_frame = tk.Frame(details_lf); area_btn_frame.grid(row=6, column=0, columnspan=4, sticky='w', pady=(5,0))
            area_text = f"Area: {step['area'][2]-step['area'][0]}x{step['area'][3]-step['area'][1]}" if step.get('area') else "Area: Global"
            w = tk.Button(area_btn_frame, text=area_text, command=self.select_area_for_step, font=('Helvetica', 9), relief=tk.FLAT); w.pack(side=tk.LEFT, padx=(0, 2))
            self.properties_widgets['area_btn'] = w
            tk.Button(area_btn_frame, text="Full Screen", command=self.set_step_area_to_fullscreen, font=('Helvetica', 9), relief=tk.FLAT).pack(side=tk.LEFT, padx=(0, 2))
            tk.Button(area_btn_frame, text="Use Global", command=self.set_step_area_to_global, font=('Helvetica', 9), relief=tk.FLAT).pack(side=tk.LEFT)
        
        if step['type'] in ['png', 'color']:
            count_details_lf = tk.LabelFrame(container, text="Count Condition", padx=5, pady=5)
            count_details_lf.grid(row=4, columnspan=3, sticky='ew', pady=(0, 10))
            self.properties_widgets['count_details_lf'] = count_details_lf
            
            tk.Label(count_details_lf, text="Expression:").grid(row=0, column=0, sticky='w', pady=5)
            w = tk.Entry(count_details_lf, width=15); w.insert(0, str(step.get('count_expression', '>= 1')))
            w.grid(row=0, column=1, columnspan=2, sticky='ew', pady=5); self.properties_widgets['count_expression'] = w
            tk.Label(count_details_lf, text="e.g., >= 5, == 1, < 3").grid(row=0, column=3, sticky='w', padx=5)

            tk.Label(count_details_lf, text="Max Cycles:").grid(row=1, column=0, sticky='w', pady=5)
            w = tk.Entry(count_details_lf, width=8)
            w.insert(0, str(step.get('count_max_cycles', 1)))
            w.grid(row=1, column=1, sticky='w', pady=5)
            self.properties_widgets['count_max_cycles'] = w
            tk.Label(count_details_lf, text="(1 = fail immediately)").grid(row=1, column=2, columnspan=2, sticky='w', padx=5)
            
            count_details_lf.columnconfigure(1, weight=1)

        elif step['type'] == 'location':
            w_btn = tk.Button(details_lf,text="Get Loc (F3)", font=('Helvetica', 9), command=lambda: self.enter_f3_mode('pick_location'), relief=tk.FLAT); w_btn.grid(row=0,column=0, pady=2); self.properties_widgets['get_loc_btn'] = w_btn
            w_lbl = tk.Label(details_lf,text=str(step.get('coords'))); w_lbl.grid(row=0,column=1,padx=10); self.properties_widgets['coords_label'] = w_lbl
            w_label = tk.Label(details_lf, text="Key to Press:"); w_label.grid(row=1, column=0, sticky='w', pady=5); self.properties_widgets['key_press_label'] = w_label
            w_entry = tk.Entry(details_lf, width=15); w_entry.insert(0, step.get('key_to_press', '')); w_entry.grid(row=1, column=1, sticky='w'); self.properties_widgets['key_press_entry'] = w_entry
            _update_location_details()
            
        def add_flow(parent, row, label, action_key, goto_key):
            tk.Label(parent,text=label).grid(row=row,column=0,sticky='w',pady=2); w=tk.StringVar(value=step.get(action_key)); self.properties_widgets[action_key]=w; tk.OptionMenu(parent,w,"Next Step","Go to Step", "Stop").grid(row=row,column=1,sticky='ew'); w=tk.Entry(parent,width=5); w.insert(0,str(step.get(goto_key))); w.grid(row=row,column=2,padx=5); self.properties_widgets[goto_key]=w
        
        add_flow(flow_lf, 0, "On Success:", 'on_success_action', 'on_success_goto_step'); tk.Label(flow_lf, text="Delay (s):").grid(row=0, column=3, padx=(10,0)); w = tk.Entry(flow_lf, width=7); w.insert(0, str(step.get('delay_after'))); w.grid(row=0, column=4); self.properties_widgets['delay_after'] = w
        
        timeout_action_label = tk.Label(flow_lf); self.properties_widgets['timeout_action_label'] = timeout_action_label
        w_action_menu=tk.StringVar(value=step.get('on_timeout_action', 'Stop')); self.properties_widgets['on_timeout_action']=w_action_menu; timeout_action_menu = tk.OptionMenu(flow_lf,w_action_menu,"Next Step","Go to Step", "Stop"); w_goto=tk.Entry(flow_lf,width=5); w_goto.insert(0,str(step.get('on_timeout_goto_step', 1))); self.properties_widgets['on_timeout_goto_step']=w_goto
        timeout_label = tk.Label(flow_lf, text="Timeout (s):"); w_timeout=tk.Entry(flow_lf, width=7); w_timeout.insert(0, str(step.get('timeout', 5.0))); self.properties_widgets['timeout'] = w_timeout
        
        is_count_mode = (step.get('type') == 'png' and step.get('action') == 'PNG Count') or \
                        (step.get('type') == 'color' and step.get('action') == 'Color Count')
        is_number_fail = step.get('logical_type') == 'Number'
        timeout_label_text = "On Fail:" if (is_count_mode or is_number_fail) else "On Timeout:"
        timeout_action_label.config(text=timeout_label_text)
        timeout_action_label.grid(row=1, column=0, sticky='w', pady=2); timeout_action_menu.grid(row=1, column=1, sticky='ew'); w_goto.grid(row=1, column=2, padx=5)
        timeout_label.grid(row=1, column=3, padx=(10,0)); w_timeout.grid(row=1, column=4)
        
        self.properties_widgets['timeout_widgets'] = [timeout_action_label, timeout_action_menu, w_goto, timeout_label, w_timeout]
        
        if step['type'] in ['png', 'color']:
            _update_details_for_action()

        is_timeout_visible = (step['type'] not in ['logical']) or (step.get('logical_type') in ['Number', 'Movement Detect'])
        if not is_timeout_visible: [w.grid_remove() for w in self.properties_widgets['timeout_widgets']]

        flow_lf.columnconfigure(1, weight=1)
        if step['type'] == 'logical': _update_logical_details_frame(flow_lf)
        
        action_btn_frm = tk.Frame(container); action_btn_frm.grid(row=7, columnspan=3, sticky='ew', pady=(15,0)); btn_pack_style = {'side': tk.LEFT, 'expand': True, 'fill': tk.X, 'padx': 2}; tk.Button(action_btn_frm, text="Apply Changes", font=('Helvetica', 9, 'bold'), command=self.apply_properties_changes, relief=tk.FLAT).pack(**btn_pack_style); tk.Button(action_btn_frm, text="Duplicate Step", font=('Helvetica', 9, 'bold'), command=self.duplicate_step, relief=tk.FLAT).pack(**btn_pack_style); tk.Button(action_btn_frm, text="Delete Step", font=('Helvetica', 9, 'bold'), command=self.remove_step, relief=tk.FLAT).pack(**btn_pack_style)
        container.columnconfigure(1, weight=1)
 
    def _update_png_preview(self, step):
        if 'png_preview' not in self.properties_widgets:
            return
        
        preview_widget = self.properties_widgets['png_preview']
        path = step.get('path')

        if not path or not os.path.exists(path):
            preview_widget.config(image='', text="No Preview Available")
            return

        try:
            with Image.open(path) as img:
                img.thumbnail((200, 100)) # Resize to max 200x100
                photo = ImageTk.PhotoImage(img)
                preview_widget.config(image=photo, text="")
                preview_widget.image = photo # Keep a reference!
        except Exception as e:
            preview_widget.config(image='', text=f"Preview Error:\n{e}")
            self.log(f"Failed to create PNG preview for {os.path.basename(path)}: {e}", "orange")
 
    def open_deal_finder_window(self):
        if hasattr(self, 'deal_finder_window') and self.deal_finder_window.winfo_exists():
            self.deal_finder_window.lift()
            return

        df_win = tk.Toplevel(self.root)
        df_win.title("GE Deal Finder")
        df_win.geometry("800x600")
        df_win.transient(self.root)
        self.deal_finder_window = df_win
        df_win.config(bg=self.current_theme['bg'])

        main_frame = ttk.Frame(df_win, padding=10); main_frame.pack(fill=tk.BOTH, expand=True)
        
        filter_frame = ttk.LabelFrame(main_frame, text="Search Filters"); filter_frame.pack(padx=5, pady=5, fill=tk.X)
        filter_vars = { 'min_profit': tk.StringVar(value="1000"), 'max_buy_price': tk.StringVar(value="10000000"), 'min_volume': tk.StringVar(value="10"), 'search_query': tk.StringVar() }
        
        ttk.Label(filter_frame, text="Min Profit:").grid(row=0, column=0, padx=5, pady=5, sticky='w'); ttk.Entry(filter_frame, textvariable=filter_vars['min_profit'], width=10).grid(row=0, column=1, padx=5, pady=5)
        ttk.Label(filter_frame, text="Max Buy Price:").grid(row=0, column=2, padx=5, pady=5, sticky='w'); ttk.Entry(filter_frame, textvariable=filter_vars['max_buy_price'], width=12).grid(row=0, column=3, padx=5, pady=5)
        ttk.Label(filter_frame, text="Min Volume (1h):").grid(row=1, column=0, padx=5, pady=5, sticky='w'); ttk.Entry(filter_frame, textvariable=filter_vars['min_volume'], width=10).grid(row=1, column=1, padx=5, pady=5)
        ttk.Label(filter_frame, text="Item Name:").grid(row=1, column=2, padx=5, pady=5, sticky='w'); ttk.Entry(filter_frame, textvariable=filter_vars['search_query'], width=12).grid(row=1, column=3, padx=5, pady=5)
        
        search_btn = ttk.Button(filter_frame, text="Search"); search_btn.grid(row=0, column=4, rowspan=2, padx=5, pady=5, sticky='nsew'); filter_frame.columnconfigure(4, weight=1)
        
        tree_frame = ttk.Frame(main_frame); tree_frame.pack(padx=5, pady=(5, 5), fill=tk.BOTH, expand=True)
        cols = ('name', 'profit', 'buy_price', 'sell_price', 'volume', 'margin'); tree = ttk.Treeview(tree_frame, columns=cols, show='headings', height=5)
        tree.heading('name', text='Item', command=lambda: self._sort_tree_column(tree, 'name', False)); tree.heading('profit', text='Profit (2% Tax)', command=lambda: self._sort_tree_column(tree, 'profit', False)); tree.heading('buy_price', text='Buy Price', command=lambda: self._sort_tree_column(tree, 'buy_price', False)); tree.heading('sell_price', text='Sell Price', command=lambda: self._sort_tree_column(tree, 'sell_price', False)); tree.heading('volume', text='Volume (1h)', command=lambda: self._sort_tree_column(tree, 'volume', False)); tree.heading('margin', text='Margin', command=lambda: self._sort_tree_column(tree, 'margin', False))
        for col, width in [('name', 200), ('profit', 100), ('buy_price', 100), ('sell_price', 100), ('volume', 100), ('margin', 100)]: tree.column(col, width=width, anchor='w')
        tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        v_scroll = ttk.Scrollbar(tree_frame, orient="vertical", command=tree.yview); v_scroll.pack(side=tk.RIGHT, fill=tk.Y); tree.configure(yscrollcommand=v_scroll.set)

        action_frame = ttk.LabelFrame(main_frame, text="Actions"); action_frame.pack(padx=5, pady=(0, 5), fill=tk.X)
        status_label = ttk.Label(action_frame, text="Ready to search."); status_label.pack(fill=tk.X, expand=True, pady=2, padx=5)

        df_btn_frame = ttk.Frame(action_frame)
        df_btn_frame.pack(fill=tk.X, expand=True, pady=(2,5), padx=5)
        df_btn_frame.columnconfigure((0,1), weight=1)
        
        use_item_btn = ttk.Button(df_btn_frame, text="Use Item in GE Interface", state=tk.DISABLED)
        use_item_btn.grid(row=0, column=0, sticky='ew', padx=(0,2))
        create_inject_step_btn = ttk.Button(df_btn_frame, text="Create GE Inject Step", state=tk.DISABLED)
        create_inject_step_btn.grid(row=0, column=1, sticky='ew', padx=(2,0))

        def on_tree_select(event):
            is_selected = True if tree.selection() else False
            use_item_btn.config(state=tk.NORMAL if is_selected else tk.DISABLED)
            create_inject_step_btn.config(state=tk.NORMAL if is_selected else tk.DISABLED)
        
        def _use_deal_in_interface():
            selected_item_id = tree.focus()
            if not selected_item_id: return
            item_data = tree.item(selected_item_id, 'values')
            name = item_data[0]
            self.ge_interface_item_name.set(name)
            self.log(f"Set GE Interface item to '{name}' from Deal Finder.")
            self.update_ge_interface_price() 
            df_win.destroy()

        def _create_ge_inject_step_from_deal():
            selected_item_id = tree.focus()
            if not selected_item_id: return
            item_data = tree.item(selected_item_id, 'values')
            name = item_data[0]
            self.add_step('logical')
            new_step = self.steps[-1]
            new_step['name'] = f"Inject: {name}"
            new_step['logical_type'] = 'GE Inject'
            new_step['ge_inject_field'] = 'Name'
            new_step['ge_inject_name'] = name
            self.log(f"Created 'GE Inject' step for '{name}' from Deal Finder.")
            self.selected_items = [{'type': 'step', 'index': len(self.steps) - 1}]
            self.populate_properties_panel()
            self.redraw_flowchart()
            df_win.destroy()

        def _search_for_deals():
            search_btn.config(state=tk.DISABLED); status_label.config(text="Fetching API data..."); df_win.update_idletasks(); tree.delete(*tree.get_children())
            
            if self.get_item_mapping() is None: status_label.config(text="Error: Could not load item map."); search_btn.config(state=tk.NORMAL); return
            all_prices = self.get_all_latest_prices(); 
            if not all_prices: status_label.config(text="Error: Could not load price data."); search_btn.config(state=tk.NORMAL); return
            all_volumes = self.get_all_hourly_volumes() 
            
            status_label.config(text="Analyzing data..."); df_win.update_idletasks()
            try:
                min_profit = int(filter_vars['min_profit'].get()); max_price = int(filter_vars['max_buy_price'].get()); min_volume = int(filter_vars['min_volume'].get()); query = filter_vars['search_query'].get().lower()
            except ValueError: status_label.config(text="Error: Invalid filter values."); search_btn.config(state=tk.NORMAL); return
            
            id_map, results = self.item_mapping_cache.get('by_id', {}), []
            for item_id, data in all_prices.items():
                if not (data and data.get('high') is not None and data.get('low') is not None and data['high'] > 0 and data['low'] > 0): continue
                buy_price, sell_price = data['low'], data['high']
                volume_data = all_volumes.get(item_id) if all_volumes else {}
                volume = (volume_data.get('highPriceVolume', 0) if volume_data else 0) + (volume_data.get('lowPriceVolume', 0) if volume_data else 0)
                item_info = id_map.get(int(item_id))
                
                if not item_info: continue
                profit = int(sell_price * 0.98) - buy_price
                
                if (buy_price < max_price and volume >= min_volume and profit >= min_profit and (query in item_info['name'].lower())):
                    results.append({'name': item_info['name'], 'profit': profit, 'buy': buy_price, 'sell': sell_price, 'volume': volume, 'margin': sell_price - buy_price})
            
            results.sort(key=lambda x: x['profit'], reverse=True)
            for res in results: tree.insert('', 'end', values=(res['name'], f"{res['profit']:,}", f"{res['buy']:,}", f"{res['sell']:,}", f"{res['volume']:,}", f"{res['margin']:,}"))
            status_label.config(text=f"Search complete. Found {len(results)} deals."); search_btn.config(state=tk.NORMAL)

        tree.bind('<<TreeviewSelect>>', on_tree_select)
        search_btn.config(command=lambda: threading.Thread(target=_search_for_deals, daemon=True).start())
        use_item_btn.config(command=_use_deal_in_interface)
        create_inject_step_btn.config(command=_create_ge_inject_step_from_deal)
        self.update_widget_colors_recursive(df_win, self.current_theme)

    def _sort_tree_column(self, tree, col, reverse):
        data_list = [(tree.set(k, col), k) for k in tree.get_children('')]
        try: data_list.sort(key=lambda t: int(str(t[0]).replace(',', '')), reverse=reverse)
        except ValueError: data_list.sort(key=lambda t: str(t[0]).lower(), reverse=reverse)
        for index, (val, k) in enumerate(data_list): tree.move(k, '', index)
        tree.heading(col, command=lambda: self._sort_tree_column(tree, col, not reverse))

    def populate_note_properties(self, container, i):
        note = self.annotations[i]
        tk.Label(container, text="Properties for Note", font=('Helvetica', 11, 'bold')).grid(row=0, columnspan=3, sticky='w', pady=(0,10)); tk.Label(container, text="Text:").grid(row=1, column=0, sticky='nw', pady=2); w = tk.Text(container, height=4, width=30, wrap=tk.WORD); w.insert('1.0', note.get('text', '')); w.grid(row=1, column=1, columnspan=2, sticky='ew'); self.properties_widgets['note_text'] = w
        tk.Label(container, text="Opacity:").grid(row=2, column=0, sticky='w', pady=2); w = tk.StringVar(value=note.get('opacity', '50%')); self.properties_widgets['note_opacity'] = w; tk.OptionMenu(container, w, "100%", "75%", "50%", "25%", "0% (Border Only)").grid(row=2, column=1, columnspan=2, sticky='w')
        tk.Label(container, text="Color:").grid(row=3, column=0, sticky='w', pady=2); color_frame = tk.Frame(container); color_frame.grid(row=3, column=1, columnspan=2, sticky='w'); w = tk.Label(color_frame, text=" ", bg=note.get('color'), relief=tk.SUNKEN, width=4); w.pack(side=tk.LEFT); self.properties_widgets['note_swatch'] = w; tk.Button(color_frame, text="Choose Color", font=('Helvetica', 9), command=self.choose_note_color, relief=tk.FLAT).pack(side=tk.LEFT, padx=5)
        action_btn_frm = tk.Frame(container); action_btn_frm.grid(row=4, columnspan=3, sticky='ew', pady=(15,0)); btn_pack_style = {'side': tk.LEFT, 'expand': True, 'fill': tk.X, 'padx': 2}; tk.Button(action_btn_frm, text="Apply Changes", font=('Helvetica', 9, 'bold'), command=self.apply_properties_changes, relief=tk.FLAT).pack(**btn_pack_style); tk.Button(action_btn_frm, text="Duplicate Note", font=('Helvetica', 9, 'bold'), command=self.duplicate_note, relief=tk.FLAT).pack(**btn_pack_style); tk.Button(action_btn_frm, text="Delete Note", font=('Helvetica', 9, 'bold'), command=self.remove_note, relief=tk.FLAT).pack(**btn_pack_style)
        container.columnconfigure(1, weight=1)

    def reset_movement_baseline(self):
        if not (self.selected_items and len(self.selected_items) == 1 and self.selected_items[0]['type'] == 'step'): return
        step = self.steps[self.selected_items[0]['index']]
        if step.get('logical_type') == 'Movement Detect': 
            step['_previous_frame_for_movement'] = None
            self.log(f"Reset movement comparison for Step {self.selected_items[0]['index'] + 1}.")
            self.populate_properties_panel()

    def choose_note_color(self):
        if not (self.selected_items and len(self.selected_items) == 1 and self.selected_items[0]['type'] == 'note'): return
        initial_color = self.annotations[self.selected_items[0]['index']].get('color'); color_code = colorchooser.askcolor(title="Choose note color", initialcolor=initial_color)
        if color_code and color_code[1]: self.properties_widgets['note_swatch'].config(bg=color_code[1])

    def apply_properties_changes(self):
        if len(self.selected_items) != 1: return
        w, item_type = self.properties_widgets, self.selected_items[0]['type']
        try:
            if item_type == 'step':
                i = self.selected_items[0]['index']; s = self.steps[i]
                s['name'] = w['name_entry'].get()
                s['enable_logging'] = w['enable_logging'].get()
                
                if 'show_area' in w:
                    s['show_area'] = w['show_area'].get()
    
                s['delay_after']=float(w['delay_after'].get()); s['on_success_action']=w['on_success_action'].get(); s['on_success_goto_step']=int(w['on_success_goto_step'].get())
                if s['type'] == 'logical':
                    s['logical_type'] = w['logical_type'].get()
                    if s['logical_type'] == 'Wait': 
                        s['max_time'] = float(w['max_time'].get())
                        s['reset_on_start'] = w['reset_on_start'].get()
                        s.pop('timeout', None); s.pop('on_timeout_action', None); s.pop('on_timeout_goto_step', None); s.pop('reset_on_timeout', None)
                    elif s['logical_type'] == 'Count': 
                        s['max_count'] = int(w['max_count'].get()); s['on_count_reached_action'] = w['on_count_reached_action'].get(); s['on_count_reached_goto_step'] = int(w['on_count_reached_goto_step'].get()); s['on_count_reached_delay'] = float(w['on_count_reached_delay'].get()); s['reset_on_start'] = w['reset_on_start'].get()
                        s['reset_on_reach'] = w['reset_on_reach'].get()
                    elif s['logical_type'] == 'Type Text': 
                        s['text_source'] = w['text_source'].get()
                        s['text_to_type'] = w['text_to_type'].get()
                        s['ge_data_field'] = w['ge_data_field'].get()
                        s['press_enter'] = w['press_enter'].get()
                        s['enter_press_delay'] = float(w['enter_press_delay'].get())
                    elif s['logical_type'] == 'GE Inject':
                        s['ge_inject_field'] = w['ge_inject_field'].get()
                        s['ge_inject_name'] = w['ge_inject_name'].get()
                        s['ge_inject_quantity'] = w['ge_inject_quantity'].get()
                        s['ge_inject_refresh'] = w['ge_inject_refresh'].get()
                    elif s['logical_type'] == 'Settings Inject':
                        s['inject_setting_name'] = w['inject_setting_name'].get()
                        s['inject_setting_value'] = w['inject_setting_value'].get()
                    elif s['logical_type'] == 'Number':
                        s['expression']=w['expression'].get(); s['timeout']=float(w['timeout'].get()); s['on_timeout_action']=w['on_timeout_action'].get(); s['on_timeout_goto_step']=int(w['on_timeout_goto_step'].get()); s['psm_mode'] = w['psm_mode'].get(); s['oem_mode'] = w['oem_mode'].get(); s['image_mode'] = w['number_image_mode'].get()
                    elif s['logical_type'] == 'Movement Detect':
                        s['movement_tolerance'] = float(w['movement_tolerance'].get())
                        s['reset_on_start'] = w['reset_on_start'].get()
                        s['timeout']=float(w['timeout'].get())
                        s['on_timeout_action']=w['on_timeout_action'].get()
                        s['on_timeout_goto_step']=int(w['on_timeout_goto_step'].get())
                else:
                    s['timeout']=float(w['timeout'].get()); s['on_timeout_action']=w['on_timeout_action'].get(); s['on_timeout_goto_step']=int(w['on_timeout_goto_step'].get()); s['action']=w['action'].get()
                    if s['type']=='color': 
                        s['tolerance']=int(w['tolerance'].get())
                        s['color_space']=w['color_space'].get()
                        s['pixel_detect_enabled'] = w['pixel_detect_enabled'].get()
                        if not s['pixel_detect_enabled']:
                            s['min_pixel_area'] = int(w['min_pixel_area'].get())
                        if s.get('action') == 'Color Count':
                            s['count_expression'] = w['count_expression'].get()
                            s['count_max_cycles'] = int(w['count_max_cycles'].get())
                    elif s['type']=='png':
                        current_path = s.get('path', '')
                        image_mode_to_clear = w['image_mode'].get()
                        cache_key = f"{current_path}|{image_mode_to_clear}"
    
                        if cache_key in self.template_cache:
                            del self.template_cache[cache_key]
                            self.log(f"Refreshed cache for: {os.path.basename(current_path)}")
    
                        if cache_key in self.folder_image_cache:
                            del self.folder_image_cache[cache_key]
                            self.log(f"Refreshed cache for folder: {os.path.basename(current_path)}")
                        
                        s['threshold']=float(w['threshold'].get())
                        s['mode']=w['png_mode'].get()
                        s['image_mode']=w['image_mode'].get()
                        s['find_first_match'] = w['find_first_match'].get()
                        if s.get('action') == 'PNG Count':
                            s['count_expression'] = w['count_expression'].get()
                            s['count_max_cycles'] = int(w['count_max_cycles'].get())
                            for old_prop in ['counter_value', 'max_count', 'reset_on_start', 'reset_on_reach', 'on_count_reached_action', 'on_count_reached_goto_step', 'on_count_reached_delay']:
                                s.pop(old_prop, None)

                    elif s['type']=='location': s['key_to_press'] = w['key_press_entry'].get()
                self.log(f"Applied changes to Step {i+1}.")
            elif item_type == 'note':
                i = self.selected_items[0]['index']; n = self.annotations[i]
                n['text'] = w['note_text'].get("1.0", tk.END).strip(); n['opacity'] = w['note_opacity'].get(); n['color'] = w['note_swatch'].cget('bg')
                self.log(f"Applied changes to Note.")
            self.redraw_flowchart()
            self.populate_properties_panel()
        except Exception as e: messagebox.showerror("Invalid Input",f"Please check your input values.\n\nDetails: {e}")

    def apply_multi_properties_changes(self):
        if len(self.selected_items) <= 1 or self.selected_items[0]['type'] != 'step':
            return

        w = self.properties_widgets
        changed_props = {}

        # Collect changed values from widgets
        for prop, widget_var in w.items():
            if prop in ['container']: continue
            try:
                value = widget_var.get()
                if value not in [self.MULTIPLE_VALUES, "", None]:
                    # Type conversion
                    if prop in ['delay_after', 'timeout', 'threshold']: value = float(value)
                    elif prop == 'tolerance': value = int(value)
                    changed_props[prop] = value
            except (tk.TclError, ValueError):
                continue # Ignore widgets that don't have a get() method or have invalid values

        if not changed_props:
            self.log("No changes to apply to multiple steps.")
            return

        # Apply changes to all selected steps
        for item in self.selected_items:
            step = self.steps[item['index']]
            for prop, value in changed_props.items():
                step[prop] = value
    
        self.log(f"Applied batch changes to {len(self.selected_items)} steps.", "green")
        self.redraw_flowchart()
        self.populate_properties_panel()


