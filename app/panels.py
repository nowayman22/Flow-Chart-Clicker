import tkinter as tk
from tkinter import ttk, scrolledtext
import os
from app import PYTESSERACT_AVAILABLE

class PanelsMixin:
    def build_ui(self):
        main_pane = ttk.PanedWindow(self.root, orient=tk.HORIZONTAL); main_pane.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        canvas_container = ttk.LabelFrame(main_pane, text="Flowchart Editor", padding=5)
        
        search_frame = ttk.Frame(canvas_container)
        search_frame.pack(side=tk.TOP, fill=tk.X, pady=(0, 5))
        ttk.Label(search_frame, text="Search:").pack(side=tk.LEFT, padx=(2, 5))
        search_entry = ttk.Entry(search_frame, textvariable=self.search_query)
        search_entry.pack(side=tk.LEFT, fill=tk.X, expand=True)
        search_btn = ttk.Button(search_frame, text="Find", command=self.search_flowchart)
        search_btn.pack(side=tk.LEFT, padx=(5, 2))
        clear_btn = ttk.Button(search_frame, text="Clear", command=self.clear_search)
        clear_btn.pack(side=tk.LEFT, padx=(0, 2))
        search_entry.bind('<Return>', self.search_flowchart)
        
        h_scroll = ttk.Scrollbar(canvas_container, orient=tk.HORIZONTAL); v_scroll = ttk.Scrollbar(canvas_container, orient=tk.VERTICAL); self.canvas = tk.Canvas(canvas_container, bg="#3c3c3c", highlightthickness=0, xscrollcommand=h_scroll.set, yscrollcommand=v_scroll.set); h_scroll.config(command=self.canvas.xview); v_scroll.config(command=self.canvas.yview); h_scroll.pack(side=tk.BOTTOM, fill=tk.X); v_scroll.pack(side=tk.RIGHT, fill=tk.Y); self.canvas.pack(fill=tk.BOTH, expand=True); main_pane.add(canvas_container, weight=3); self.canvas.bind("<ButtonPress-1>", self.on_canvas_press); self.canvas.bind("<B1-Motion>", self.on_drag_motion); self.canvas.bind("<ButtonRelease-1>", self.on_drag_release); self._bind_mouse_scroll()
        right_panel = ttk.Frame(main_pane); right_panel.pack(fill=tk.Y); main_pane.add(right_panel, weight=1)
        
        status_display_frame = ttk.Frame(right_panel, padding=5); status_display_frame.pack(fill=tk.X, pady=(5, 5))
        header_frame = ttk.Frame(status_display_frame); header_frame.pack(fill=tk.X)
        self.status_label = ttk.Label(header_frame, text="Status: Stopped", anchor="w", font=('Helvetica', 11, 'bold')); self.status_label.pack(side=tk.LEFT, fill=tk.X, expand=True)
        self.cycle_time_label = ttk.Label(header_frame, textvariable=self.cycle_time_display, anchor='e', font=('Helvetica', 10)); self.cycle_time_label.pack(side=tk.RIGHT)
        self.timeout_countdown_label = ttk.Label(status_display_frame, text="", anchor='w', font=('Helvetica', 9, 'italic')); self.timeout_countdown_label.pack(fill=tk.X)
        self.delay_countdown_label = ttk.Label(status_display_frame, text="", anchor='w', font=('Helvetica', 9, 'italic')); self.delay_countdown_label.pack(fill=tk.X)

        detection_info_frame = ttk.LabelFrame(right_panel, text="Live Detection Info", padding=5); detection_info_frame.pack(fill=tk.X, padx=5, pady=(0, 10)); ttk.Label(detection_info_frame, textvariable=self.last_detection_info, anchor='w', font=('Consolas', 10)).pack(fill=tk.X)

        # --- EDIT START: Implement vertical PanedWindow for resizable tabs/buttons ---
        # 1. Create a new vertical PanedWindow inside the main right_panel
        right_sub_pane = ttk.PanedWindow(right_panel, orient=tk.VERTICAL)
        right_sub_pane.pack(fill=tk.BOTH, expand=True, padx=5, pady=(0, 5))

        # 2. Change the parent of the notebook to the new sub-pane. Do NOT .pack() it.
        notebook = ttk.Notebook(right_sub_pane)
        self.props_tab = ttk.Frame(notebook, padding=10); globals_tab = ttk.Frame(notebook, padding=10); testing_tab = ttk.Frame(notebook, padding=10); log_tab = ttk.Frame(notebook, padding=10); ge_interface_tab = ttk.Frame(notebook, padding=10); info_tab = ttk.Frame(notebook, padding=10)
        notebook.add(self.props_tab, text='Properties'); notebook.add(ge_interface_tab, text='GE Interface'); notebook.add(globals_tab, text='Global Settings'); notebook.add(testing_tab, text='Testing Logic'); notebook.add(log_tab, text='Execution Log'); notebook.add(info_tab, text='Info')
        self.build_properties_panel(); self.build_ge_interface_panel(ge_interface_tab); self.build_globals_panel(globals_tab); self.build_testing_panel(testing_tab); self.build_log_panel(log_tab); self.build_info_panel(info_tab)
        
        # 3. Change the parent of the control_bar to the new sub-pane. Do NOT .pack() it.
        control_bar = ttk.Frame(right_sub_pane, padding=5)

        # 4. Add the notebook and control bar to the pane using .add()
        right_sub_pane.add(notebook, weight=1)      # Give notebook more space by default
        right_sub_pane.add(control_bar, weight=0)   # Give control bar minimum space
        # --- EDIT END ---

        btn_style = {'fg': 'white', 'height': 2, 'font': ('Helvetica', 9, 'bold'), 'relief': tk.FLAT, 'borderwidth': 0}

        add_step_grid = ttk.Frame(control_bar); add_step_grid.pack(fill=tk.X, pady=(0, 5))
        add_step_grid.columnconfigure((0, 1), weight=1)

        grid_btn_style = {'sticky': 'ew', 'padx': 2, 'pady': (0, 3)}
        tk.Button(add_step_grid, text="+ Color Step", **btn_style, command=lambda: self.add_step('color')).grid(row=0, column=0, **grid_btn_style)
        tk.Button(add_step_grid, text="+ PNG Step", **btn_style, command=lambda: self.add_step('png')).grid(row=0, column=1, **grid_btn_style)
        tk.Button(add_step_grid, text="+ Click / Press Step", **btn_style, command=lambda: self.add_step('location')).grid(row=1, column=0, **grid_btn_style)
        tk.Button(add_step_grid, text="+ Logical Step", **btn_style, command=lambda: self.add_step('logical')).grid(row=1, column=1, **grid_btn_style)
        tk.Button(add_step_grid, text="+ Add Note", **btn_style, command=self.add_annotation).grid(row=2, column=0, columnspan=2, **grid_btn_style)
        
        start_frm = ttk.Frame(control_bar); start_frm.pack(fill=tk.X, pady=2); self.start_btn = ttk.Button(start_frm, text="Start (F2)", command=self.start); self.start_btn.pack(side=tk.LEFT, expand=True, fill=tk.BOTH, padx=(0, 5), ipady=5); self.stop_btn = ttk.Button(start_frm, text="Stop (F2)", command=self.stop, state=tk.DISABLED); self.stop_btn.pack(side=tk.LEFT, expand=True, fill=tk.BOTH, ipady=5)
        start_at_frm = ttk.Frame(start_frm); start_at_frm.pack(side=tk.LEFT, fill=tk.NONE, padx=(10, 0)); ttk.Label(start_at_frm, text="Start Step:").pack(anchor='s'); ttk.Entry(start_at_frm, textvariable=self.start_step, width=5).pack(anchor='n', pady=(2,0))
        file_ops_frm = ttk.Frame(control_bar); file_ops_frm.pack(fill=tk.X, pady=(5, 0)); btn_pack_style = {'side': tk.LEFT, 'expand': True, 'fill': tk.X, 'padx': 2}; tk.Button(file_ops_frm, text="Import JSON", **btn_style, command=self.import_from_json).pack(**btn_pack_style); tk.Button(file_ops_frm, text="Export JSON", **btn_style, command=self.export_to_json).pack(**btn_pack_style); tk.Button(file_ops_frm, text="Reset All", **btn_style, command=self.reset_all).pack(**btn_pack_style)
   
    def build_properties_panel(self):
        self.properties_widgets['default_label'] = ttk.Label(self.props_tab, text="\n\nSelect a step or note in the flowchart\nto view and edit its properties.", justify=tk.CENTER, font=('Helvetica', 10)); self.properties_widgets['default_label'].pack(expand=True, fill=tk.BOTH)
        if self.current_theme: self.properties_widgets['default_label'].config(foreground=self.current_theme['node_text_grey'])

    def build_globals_panel(self, parent):
        self._sync_global_settings_ui_from_model()
        parent.columnconfigure(0, weight=1) # Make the single column expandable

        # --- Mouse Movement Section ---
        mouse_lf = ttk.LabelFrame(parent, text="Mouse Movement")
        mouse_lf.grid(row=0, column=0, sticky='ew', pady=(0, 10), padx=2)
        mouse_lf.columnconfigure(1, weight=1)
        
        ttk.Label(mouse_lf, text="Mode:").grid(row=0, column=0, sticky='w', padx=5, pady=5)
        
        mode_radio_frame = ttk.Frame(mouse_lf)
        mode_radio_frame.grid(row=0, column=1, sticky='ew', padx=5, pady=5)
        self.mouse_move_mode.trace_add("write", self._update_mouse_mode_visibility)
        ttk.Radiobutton(mode_radio_frame, text="Regular", variable=self.mouse_move_mode, value="Regular").grid(row=0, column=0, padx=(0, 5))
        ttk.Radiobutton(mode_radio_frame, text="Dynamic", variable=self.mouse_move_mode, value="Dynamic").grid(row=0, column=1, padx=5)
        ttk.Radiobutton(mode_radio_frame, text="Pixels/Sec", variable=self.mouse_move_mode, value="Pixels Per Second").grid(row=0, column=2, padx=5)
        
        # Frame for Regular mode settings
        self.regular_speed_frame = ttk.Frame(mouse_lf)
        self.regular_speed_frame.grid(row=1, column=0, columnspan=2, sticky='ew', padx=5)
        self.regular_speed_frame.columnconfigure(1, weight=1)
        ttk.Label(self.regular_speed_frame, text="Base Mouse Speed (s):").grid(row=0, column=0, sticky="w", pady=2)
        ttk.Entry(self.regular_speed_frame, textvariable=self.global_settings_ui_vars['mouse_speed'], width=10).grid(row=0, column=1, sticky="ew", pady=2)

        # Frame for Dynamic mode settings
        self.dynamic_speed_frame = ttk.Frame(mouse_lf)
        self.dynamic_speed_frame.grid(row=2, column=0, columnspan=2, sticky='ew', padx=5)
        self.dynamic_speed_frame.columnconfigure(1, weight=1)
        ttk.Label(self.dynamic_speed_frame, text="Min Move Time (s):").grid(row=0, column=0, sticky="w", pady=2)
        ttk.Entry(self.dynamic_speed_frame, textvariable=self.global_settings_ui_vars['min_move_time'], width=10).grid(row=0, column=1, sticky="ew", pady=2)
        ttk.Label(self.dynamic_speed_frame, text="Max Move Time (s):").grid(row=1, column=0, sticky="w", pady=2)
        ttk.Entry(self.dynamic_speed_frame, textvariable=self.global_settings_ui_vars['max_move_time'], width=10).grid(row=1, column=1, sticky="ew", pady=2)

        # Frame for Pixels Per Second mode settings
        self.pps_speed_frame = ttk.Frame(mouse_lf)
        self.pps_speed_frame.grid(row=3, column=0, columnspan=2, sticky='ew', padx=5)
        self.pps_speed_frame.columnconfigure(1, weight=1)
        ttk.Label(self.pps_speed_frame, text="Pixels Per Second:").grid(row=0, column=0, sticky="w", pady=2)
        ttk.Entry(self.pps_speed_frame, textvariable=self.global_settings_ui_vars['pixels_per_second'], width=10).grid(row=0, column=1, sticky="ew", pady=2)

        self._update_mouse_mode_visibility()

        # --- Click Variation Section ---
        variation_lf = ttk.LabelFrame(parent, text="Click Variation")
        variation_lf.grid(row=1, column=0, sticky='ew', pady=(0, 10), padx=2)
        variation_lf.columnconfigure(1, weight=1)
        ttk.Label(variation_lf, text="Location Offset (±px):").grid(row=0, column=0, sticky="w", pady=2, padx=5)
        ttk.Entry(variation_lf, textvariable=self.global_settings_ui_vars['loc_offset_variance'], width=10).grid(row=0, column=1, sticky="ew", pady=2, padx=5)
        ttk.Label(variation_lf, text="Speed Variance (±s):").grid(row=1, column=0, sticky="w", pady=2, padx=5)
        ttk.Entry(variation_lf, textvariable=self.global_settings_ui_vars['speed_variance'], width=10).grid(row=1, column=1, sticky="ew", pady=2, padx=5)
        ttk.Label(variation_lf, text="Hold Variance (±s):").grid(row=2, column=0, sticky="w", pady=2, padx=5)
        ttk.Entry(variation_lf, textvariable=self.global_settings_ui_vars['hold_duration_variance'], width=10).grid(row=2, column=1, sticky="ew", pady=2, padx=5)
        
        # --- Global Timings Section ---
        timing_lf = ttk.LabelFrame(parent, text="Global Timings")
        timing_lf.grid(row=2, column=0, sticky='ew', pady=(0, 10), padx=2)
        timing_lf.columnconfigure(1, weight=1)
        ttk.Label(timing_lf, text="Scan Interval (s):").grid(row=0, column=0, sticky="w", pady=2, padx=5)
        ttk.Entry(timing_lf, textvariable=self.global_settings_ui_vars['scan_interval'], width=10).grid(row=0, column=1, sticky="ew", pady=2, padx=5)
        ttk.Label(timing_lf, text="Base Hold Duration (s):").grid(row=1, column=0, sticky="w", pady=2, padx=5)
        ttk.Entry(timing_lf, textvariable=self.global_settings_ui_vars['hold_duration'], width=10).grid(row=1, column=1, sticky="ew", pady=2, padx=5)

        # --- Flowchart Grid Section ---
        flowchart_lf = ttk.LabelFrame(parent, text="Flowchart Grid")
        flowchart_lf.grid(row=3, column=0, sticky='ew', pady=(0, 10), padx=2)
        flowchart_lf.columnconfigure(1, weight=1)
        ttk.Checkbutton(flowchart_lf, text="Show Grid", variable=self.grid_visible, command=self.redraw_flowchart).grid(row=0, column=0, columnspan=2, sticky='w', pady=2, padx=5)
        ttk.Checkbutton(flowchart_lf, text="Enable Grid Latching", variable=self.grid_latching).grid(row=1, column=0, columnspan=2, sticky='w', pady=2, padx=5)
        ttk.Label(flowchart_lf, text="Grid Spacing (px):").grid(row=2, column=0, sticky="w", pady=2, padx=5)
        ttk.Entry(flowchart_lf, textvariable=self.global_settings_ui_vars['grid_spacing'], width=10).grid(row=2, column=1, sticky="ew", pady=2, padx=5)
        ttk.Label(flowchart_lf, text="Grid Opacity:").grid(row=3, column=0, sticky="w", pady=2, padx=5)
        ttk.Scale(flowchart_lf, from_=0.0, to=1.0, orient=tk.HORIZONTAL, variable=self.grid_opacity, command=lambda e: self.redraw_flowchart()).grid(row=3, column=1, sticky="ew", pady=2, padx=5)
        
        # --- Global Area Section ---
        garea_lf = ttk.LabelFrame(parent, text="Global Area (F4)")
        garea_lf.grid(row=4, column=0, sticky='ew', pady=(0, 10), padx=2)
        garea_lf.columnconfigure(0, weight=1)
        area_frame=ttk.Frame(garea_lf)
        area_frame.grid(row=0,column=0,sticky="ew")
        area_vars = [self.global_settings_ui_vars['area_x1'], self.global_settings_ui_vars['area_y1'], self.global_settings_ui_vars['area_x2'], self.global_settings_ui_vars['area_y2']]
        for i, var in enumerate(area_vars):
            entry = ttk.Entry(area_frame, textvariable=var, width=7)
            entry.grid(row=0, column=i, padx=2, sticky='ew')
            area_frame.columnconfigure(i, weight=1)
        btn_frame = ttk.Frame(garea_lf)
        btn_frame.grid(row=1, column=0, sticky="ew", pady=(5,0))
        btn_frame.columnconfigure((0,1), weight=1)
        tk.Button(btn_frame, text="Select Area (F4)", font=('Helvetica', 9, 'bold'), command=self.select_area_mode, relief=tk.FLAT).grid(row=0, column=0, sticky='ew', padx=(0,1))
        tk.Button(btn_frame, text="Full Screen", font=('Helvetica', 9, 'bold'), command=self.set_global_area_to_fullscreen, relief=tk.FLAT).grid(row=0, column=1, sticky='ew', padx=(1,0))

        # --- Convenience Section ---
        convenience_lf = ttk.LabelFrame(parent, text="Convenience")
        convenience_lf.grid(row=5, column=0, sticky='ew', pady=(0, 10), padx=2)
        ttk.Checkbutton(convenience_lf, text="Hide Window on Capture (F3/F4)", variable=self.hide_on_select, onvalue=True, offvalue=False).grid(row=0, column=0, sticky='w', pady=1, padx=5)
        ttk.Checkbutton(convenience_lf, text="Enable all 'Show Area' overlays", variable=self.enable_all_show_area, command=self.update_all_area_overlays).grid(row=1, column=0, sticky='w', pady=1, padx=5)
        ttk.Checkbutton(convenience_lf, text="Update 'Start Step' when stopped", variable=self.start_at_stopped_pos).grid(row=2, column=0, sticky='w', pady=1, padx=5)
        
        # --- Apply Button ---
        tk.Button(parent, text="Apply Global Settings", font=('Helvetica', 10, 'bold'), command=self.apply_global_settings, relief=tk.FLAT).grid(row=6, column=0, sticky='ew', pady=(5,5), ipady=4)

    def build_log_panel(self, parent):
        log_controls_frame = ttk.Frame(parent); log_controls_frame.pack(fill=tk.X, pady=(0, 5))
        
        ttk.Label(log_controls_frame, text="Search:").pack(side=tk.LEFT, padx=(0, 5))
        ttk.Entry(log_controls_frame, textvariable=self.log_search_query, width=20).pack(side=tk.LEFT, expand=True, fill=tk.X)
        
        tk.Button(log_controls_frame, text="Clear", command=self.clear_log, font=('Helvetica', 9), relief=tk.FLAT).pack(side=tk.LEFT, padx=(10, 0))
        ttk.Label(log_controls_frame, text="Auto-clear (lines):").pack(side=tk.LEFT, padx=(10, 5))
        ttk.Entry(log_controls_frame, textvariable=self.log_auto_clear_lines, width=6).pack(side=tk.LEFT)
        
        self.log_text = scrolledtext.ScrolledText(parent,state='disabled',wrap=tk.WORD,borderwidth=0,highlightthickness=1); self.log_text.pack(fill=tk.BOTH,expand=True)
        if self.current_theme: self.log_text.config(highlightbackground=self.current_theme['node_border'], highlightcolor=self.current_theme['node_border'])

    def build_info_panel(self, parent):
        info_frame = ttk.Frame(parent, padding=10)
        info_frame.pack(fill=tk.BOTH, expand=True)

        info_frame.columnconfigure(1, weight=1)

        header_font = ('Helvetica', 11, 'bold')
        label_font = ('Helvetica', 10, 'bold')
        value_font = ('Helvetica', 10)
        
        ttk.Label(info_frame, text="Application Information", font=header_font).grid(row=0, column=0, columnspan=2, pady=(0, 15), sticky='w')

        info_data = {
            "Application:": "Flowchart Automation Tool",
            "Version #:": "Rev. 66",
            "Description:": "A visual, node-based automation tool for creating and running complex task sequences using image, color, and text (OCR) detection.",
            "Last Updated:": "October 1, 2025",
            "Creator:": "u/WonRu2",
            "Key Dependencies:": "Tkinter, OpenCV-Python, PyAutoGUI, Keyboard, Pytesseract (Tesseract-OCR), Pillow",
            "License:": "Freeware"
        }
        
        row_num = 1
        for label_text, value_text in info_data.items():
            ttk.Label(info_frame, text=label_text, font=label_font).grid(row=row_num, column=0, sticky='nw', padx=(0, 10), pady=4)
            ttk.Label(info_frame, text=value_text, font=value_font, wraplength=400).grid(row=row_num, column=1, sticky='w', pady=4)
            row_num += 1

    def build_testing_panel(self, parent):
        self.test_area_buttons = []
        type_frame = ttk.Frame(parent); type_frame.pack(fill=tk.X, pady=(0, 10)); ttk.Label(type_frame, text="Test Type:", font=('Helvetica', 10, 'bold')).pack(side=tk.LEFT, padx=(0, 10)); ttk.Radiobutton(type_frame, text="PNG", variable=self.active_test_type, value="PNG", command=self._update_test_panel_visibility).pack(side=tk.LEFT); ttk.Radiobutton(type_frame, text="Color", variable=self.active_test_type, value="Color", command=self._update_test_panel_visibility).pack(side=tk.LEFT, padx=10)
        ttk.Radiobutton(type_frame, text="PNG Count", variable=self.active_test_type, value="PNG Count", command=self._update_test_panel_visibility).pack(side=tk.LEFT)
        ttk.Radiobutton(type_frame, text="Color Count", variable=self.active_test_type, value="Color Count", command=self._update_test_panel_visibility).pack(side=tk.LEFT, padx=10)
        if PYTESSERACT_AVAILABLE: ttk.Radiobutton(type_frame, text="Number", variable=self.active_test_type, value="Number", command=self._update_test_panel_visibility).pack(side=tk.LEFT, padx=10)
        
        self.png_test_frame = ttk.LabelFrame(parent, text="PNG Test Settings")
        png_mode_frm = ttk.Frame(self.png_test_frame); png_mode_frm.grid(row=0, columnspan=3, sticky='w', pady=(0,5)); tk.Radiobutton(png_mode_frm, text="File", variable=self.test_png_mode, value='file').pack(side=tk.LEFT); tk.Radiobutton(png_mode_frm, text="Folder", variable=self.test_png_mode, value='folder').pack(side=tk.LEFT, padx=10)
        path_frm = ttk.Frame(self.png_test_frame); path_frm.grid(row=1, columnspan=3, sticky='ew', pady=5)
        tk.Button(path_frm, text="Snip", command=self.snip_image_for_test, font=('Helvetica', 9), relief=tk.FLAT).pack(side=tk.LEFT, padx=(0,5))
        tk.Button(path_frm, text="Browse", command=self.browse_for_test_path, font=('Helvetica', 9), relief=tk.FLAT).pack(side=tk.LEFT)
        self.test_path_label = ttk.Label(path_frm, textvariable=self.test_png_path_display, anchor='w', wraplength=200, justify='left'); self.test_path_label.pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
        ttk.Label(self.png_test_frame, text="Threshold:").grid(row=2, column=0, sticky='w', pady=5); ttk.Entry(self.png_test_frame, textvariable=self.test_png_threshold, width=10).grid(row=2, column=1, sticky='w');
        
        ttk.Label(self.png_test_frame, text="Image Mode:").grid(row=3, column=0, sticky='w', pady=5);
        ttk.OptionMenu(self.png_test_frame, self.test_png_image_mode, 'Grayscale', 'Grayscale', 'Color', 'Binary (B&W)').grid(row=3, column=1, sticky='ew')
        
        self.test_png_area_btn = tk.Button(self.png_test_frame, text="Select Area", command=self.select_area_for_test, font=('Helvetica', 9), relief=tk.FLAT); self.test_png_area_btn.grid(row=2, column=2, rowspan=2, sticky='e', padx=5); self.png_test_frame.columnconfigure(2, weight=1)
        self.test_area_buttons.append(self.test_png_area_btn)

        self.color_test_frame = ttk.LabelFrame(parent, text="Color Test Settings")
        tk.Button(self.color_test_frame, text="Pick Color (F3)", font=('Helvetica', 9), command=lambda: self.enter_f3_mode('pick_test_color'), relief=tk.FLAT).grid(row=0, column=0, sticky='w', pady=2); self.test_color_swatch = tk.Label(self.color_test_frame, text=" ", bg=self.rgb_to_hex(self.test_color_rgb), relief=tk.SUNKEN, width=3); self.test_color_swatch.grid(row=0, column=1, padx=4)
        ttk.Label(self.color_test_frame, text="Tolerance:").grid(row=0, column=2, sticky='w', padx=(10,2)); ttk.Entry(self.color_test_frame, textvariable=self.test_color_tolerance, width=10).grid(row=0, column=3, sticky='w')
        ttk.Label(self.color_test_frame, text="Color Space:").grid(row=1, column=0, columnspan=2, sticky='w', pady=5)
        ttk.OptionMenu(self.color_test_frame, self.test_color_space, 'HSV', 'HSV', 'RGB').grid(row=1, column=2, columnspan=2, sticky='ew')
        self.test_color_area_btn = tk.Button(self.color_test_frame, text="Select Area", command=self.select_area_for_test, font=('Helvetica', 9), relief=tk.FLAT); self.test_color_area_btn.grid(row=0, column=4, rowspan=2, sticky='e', padx=5); self.color_test_frame.columnconfigure(4, weight=1)
        self.test_area_buttons.append(self.test_color_area_btn)

        self.color_count_test_frame = ttk.LabelFrame(parent, text="Color Count Test Settings")
        tk.Button(self.color_count_test_frame, text="Pick Color (F3)", font=('Helvetica', 9), command=lambda: self.enter_f3_mode('pick_test_color'), relief=tk.FLAT).grid(row=0, column=0, sticky='w', pady=2)
        self.test_color_swatch_count = tk.Label(self.color_count_test_frame, text=" ", bg=self.rgb_to_hex(self.test_color_rgb), relief=tk.SUNKEN, width=3); self.test_color_swatch_count.grid(row=0, column=1, padx=4)
        ttk.Label(self.color_count_test_frame, text="Tolerance:").grid(row=0, column=2, sticky='w', padx=(10,2)); ttk.Entry(self.color_count_test_frame, textvariable=self.test_color_tolerance, width=10).grid(row=0, column=3, sticky='w')
        ttk.Label(self.color_count_test_frame, text="Min Area (px):").grid(row=1, column=0, sticky='w', pady=2); ttk.Entry(self.color_count_test_frame, textvariable=self.test_color_min_area, width=10).grid(row=1, column=1, sticky='w', padx=2)
        ttk.Label(self.color_count_test_frame, text="Color Space:").grid(row=1, column=2, sticky='w', padx=(10,2)); ttk.OptionMenu(self.color_count_test_frame, self.test_color_space, 'HSV', 'HSV', 'RGB').grid(row=1, column=3, sticky='ew')
        ttk.Label(self.color_count_test_frame, text="Expression:").grid(row=2, column=0, sticky='w', pady=5); ttk.Entry(self.color_count_test_frame, textvariable=self.test_color_count_expression, width=15).grid(row=2, column=1, columnspan=3, sticky='ew', pady=5)
        self.test_color_count_area_btn = tk.Button(self.color_count_test_frame, text="Select Area", command=self.select_area_for_test, font=('Helvetica', 9), relief=tk.FLAT); self.test_color_count_area_btn.grid(row=0, column=4, rowspan=3, sticky='e', padx=5); self.color_count_test_frame.columnconfigure(4, weight=1)
        self.test_area_buttons.append(self.test_color_count_area_btn)

        self.number_test_frame = ttk.LabelFrame(parent, text="Number Test Settings")
        ttk.Label(self.number_test_frame, text="Expression:").grid(row=0, column=0, sticky='w', pady=5); ttk.Entry(self.number_test_frame, textvariable=self.test_number_expression, width=15).grid(row=0, column=1, columnspan=3, sticky='ew', pady=5)
        ttk.Label(self.number_test_frame, text="OEM:").grid(row=1, column=0, sticky='w', pady=5)
        ttk.OptionMenu(self.number_test_frame, self.test_number_oem, self.test_number_oem.get(), *self.oem_options.keys()).grid(row=1, column=1, columnspan=3, sticky='ew')
        ttk.Label(self.number_test_frame, text="PSM:").grid(row=2, column=0, sticky='w', pady=5)
        ttk.OptionMenu(self.number_test_frame, self.test_number_psm, self.test_number_psm.get(), *self.psm_options.keys()).grid(row=2, column=1, columnspan=3, sticky='ew')
        self.test_number_area_btn = tk.Button(self.number_test_frame, text="Select Area", command=self.select_area_for_test, font=('Helvetica', 9), relief=tk.FLAT); self.test_number_area_btn.grid(row=0, column=4, rowspan=3, sticky='e', padx=5)
        self.number_test_frame.columnconfigure(4, weight=1)
        self.test_area_buttons.append(self.test_number_area_btn)

        self.png_count_test_frame = ttk.LabelFrame(parent, text="PNG Count Test Settings")
        png_count_mode_frm = ttk.Frame(self.png_count_test_frame); png_count_mode_frm.grid(row=0, columnspan=3, sticky='w', pady=(0,5)); tk.Radiobutton(png_count_mode_frm, text="File", variable=self.test_png_mode, value='file').pack(side=tk.LEFT); tk.Radiobutton(png_count_mode_frm, text="Folder", variable=self.test_png_mode, value='folder').pack(side=tk.LEFT, padx=10)
        path_frm_count = ttk.Frame(self.png_count_test_frame); path_frm_count.grid(row=1, columnspan=3, sticky='ew', pady=5)
        tk.Button(path_frm_count, text="Snip", command=self.snip_image_for_test, font=('Helvetica', 9), relief=tk.FLAT).pack(side=tk.LEFT, padx=(0,5))
        tk.Button(path_frm_count, text="Browse", command=self.browse_for_test_path, font=('Helvetica', 9), relief=tk.FLAT).pack(side=tk.LEFT)
        ttk.Label(path_frm_count, textvariable=self.test_png_path_display, anchor='w', wraplength=200, justify='left').pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
        ttk.Label(self.png_count_test_frame, text="Threshold:").grid(row=2, column=0, sticky='w', pady=5); ttk.Entry(self.png_count_test_frame, textvariable=self.test_png_threshold, width=10).grid(row=2, column=1, sticky='w');
        ttk.Label(self.png_count_test_frame, text="Image Mode:").grid(row=3, column=0, sticky='w', pady=5);
        ttk.OptionMenu(self.png_count_test_frame, self.test_png_image_mode, 'Grayscale', 'Grayscale', 'Color', 'Binary (B&W)').grid(row=3, column=1, sticky='ew')
        ttk.Label(self.png_count_test_frame, text="Expression:").grid(row=4, column=0, sticky='w', pady=5); ttk.Entry(self.png_count_test_frame, textvariable=self.test_png_count_expression, width=15).grid(row=4, column=1, sticky='ew', pady=5)
        self.test_png_count_area_btn = tk.Button(self.png_count_test_frame, text="Select Area", command=self.select_area_for_test, font=('Helvetica', 9), relief=tk.FLAT); self.test_png_count_area_btn.grid(row=2, column=2, rowspan=3, sticky='e', padx=5); self.png_count_test_frame.columnconfigure(2, weight=1)
        self.test_area_buttons.append(self.test_png_count_area_btn)

        tk.Button(parent, text="Run Test", font=('Helvetica', 10, 'bold'), command=self.run_test, relief=tk.FLAT).pack(fill=tk.X, pady=(15, 5), ipady=4); self.test_results_text = scrolledtext.ScrolledText(parent, state='disabled', wrap=tk.WORD, height=8, borderwidth=0, highlightthickness=1); self.test_results_text.pack(fill=tk.BOTH, expand=True, pady=(5,0))
        if self.current_theme: self.test_results_text.config(highlightbackground=self.current_theme['node_border'], highlightcolor=self.current_theme['node_border'])
        self._update_test_panel_visibility()

    def _update_test_panel_visibility(self):
        self.png_test_frame.pack_forget(); self.color_test_frame.pack_forget(); self.number_test_frame.pack_forget(); self.png_count_test_frame.pack_forget(); self.color_count_test_frame.pack_forget()
        active_type = self.active_test_type.get()
        if active_type == "PNG": self.png_test_frame.pack(fill=tk.X, pady=5)
        elif active_type == "Color": self.color_test_frame.pack(fill=tk.X, pady=5)
        elif active_type == "PNG Count": self.png_count_test_frame.pack(fill=tk.X, pady=5)
        elif active_type == "Color Count": self.color_count_test_frame.pack(fill=tk.X, pady=5)
        elif active_type == "Number": self.number_test_frame.pack(fill=tk.X, pady=5)

    def build_ge_interface_panel(self, parent):
        # --- Item Input ---
        input_lf = ttk.LabelFrame(parent, text="Item & Quantity")
        input_lf.pack(fill=tk.X, padx=5, pady=5)
        input_lf.columnconfigure(1, weight=1)
        ttk.Label(input_lf, text="Item Name:").grid(row=0, column=0, sticky='w', padx=5, pady=5)
        ttk.Entry(input_lf, textvariable=self.ge_interface_item_name).grid(row=0, column=1, sticky='ew', padx=5, pady=5)
        ttk.Label(input_lf, text="Quantity:").grid(row=1, column=0, sticky='w', padx=5, pady=5)
        ttk.Entry(input_lf, textvariable=self.ge_interface_item_quantity).grid(row=1, column=1, sticky='ew', padx=5, pady=5)
        
        # --- Main frame for strategies ---
        strategy_frame = ttk.Frame(parent)
        strategy_frame.pack(fill=tk.X, padx=5, pady=5)
        strategy_frame.columnconfigure((0, 1), weight=1)

        # --- Buy Strategy ---
        buy_strategy_lf = ttk.LabelFrame(strategy_frame, text="Buy Strategy")
        buy_strategy_lf.grid(row=0, column=0, sticky='nsew', padx=(0, 2))
        buy_strategy_options = ["Insta-Buy", "+5%", "-5%", "Custom Price", "Flip-Buy (use Insta-Sell)", "Flip-Buy (Insta-Sell + Margin)"]
        buy_om = ttk.OptionMenu(buy_strategy_lf, self.ge_interface_buy_price_strategy, self.ge_interface_buy_price_strategy.get(), *buy_strategy_options, command=self._toggle_ge_buy_options)
        buy_om.pack(fill=tk.X, padx=5, pady=5)
        self.ge_buy_custom_price_entry = ttk.Entry(buy_strategy_lf, textvariable=self.ge_interface_buy_custom_price, width=15)
        self.ge_buy_margin_entry = ttk.Entry(buy_strategy_lf, textvariable=self.ge_interface_buy_price_margin, width=15)

        # --- Sell Strategy ---
        sell_strategy_lf = ttk.LabelFrame(strategy_frame, text="Sell Strategy")
        sell_strategy_lf.grid(row=0, column=1, sticky='nsew', padx=(2, 0))
        sell_strategy_options = ["Insta-Sell", "+5%", "-5%", "Custom Price", "Flip-Sell (use Insta-Buy)", "Flip-Sell (Insta-Buy - Margin)"]
        sell_om = ttk.OptionMenu(sell_strategy_lf, self.ge_interface_sell_price_strategy, self.ge_interface_sell_price_strategy.get(), *sell_strategy_options, command=self._toggle_ge_sell_options)
        sell_om.pack(fill=tk.X, padx=5, pady=5)
        self.ge_sell_custom_price_entry = ttk.Entry(sell_strategy_lf, textvariable=self.ge_interface_sell_custom_price, width=15)
        self.ge_sell_margin_entry = ttk.Entry(sell_strategy_lf, textvariable=self.ge_interface_sell_price_margin, width=15)

        self._toggle_ge_buy_options()
        self._toggle_ge_sell_options()

        # --- Auto-Update ---
        auto_update_lf = ttk.LabelFrame(parent, text="Auto-Update")
        auto_update_lf.pack(fill=tk.X, padx=5, pady=5)
        ttk.Checkbutton(auto_update_lf, text="Enable Auto-Update", variable=self.ge_auto_update_enabled).pack(side=tk.LEFT, padx=5)
        ttk.Label(auto_update_lf, text="Interval (s):").pack(side=tk.LEFT, padx=(10, 2))
        ttk.Entry(auto_update_lf, textvariable=self.ge_auto_update_interval, width=8).pack(side=tk.LEFT, padx=2)

        # --- Actions & Results ---
        results_lf = ttk.LabelFrame(parent, text="Actions & Results")
        results_lf.pack(fill=tk.X, padx=5, pady=5)
        results_lf.columnconfigure(1, weight=1)
        
        tk.Button(results_lf, text="Find Profitable Flips...", font=('Helvetica', 9, 'bold'), command=self.open_deal_finder_window, relief=tk.FLAT).grid(row=0, column=0, columnspan=2, sticky='ew', padx=5, pady=(5,0))
        tk.Button(results_lf, text="Fetch & Calculate Price", font=('Helvetica', 10, 'bold'), command=self.update_ge_interface_price, relief=tk.FLAT).grid(row=1, column=0, columnspan=2, sticky='ew', padx=5, pady=5, ipady=4)
        
        ttk.Label(results_lf, text="Calculated Buy Price:").grid(row=2, column=0, sticky='w', padx=5, pady=2)
        ttk.Label(results_lf, textvariable=self.ge_interface_display_buy_price, font=('Consolas', 10, 'bold')).grid(row=2, column=1, sticky='w', padx=5, pady=2)
        ttk.Label(results_lf, text="Total Buy Value:").grid(row=3, column=0, sticky='w', padx=5, pady=2)
        ttk.Label(results_lf, textvariable=self.ge_interface_display_buy_total, font=('Consolas', 10, 'bold')).grid(row=3, column=1, sticky='w', padx=5, pady=2)
        ttk.Label(results_lf, text="Calculated Sell Price:").grid(row=4, column=0, sticky='w', padx=5, pady=2)
        ttk.Label(results_lf, textvariable=self.ge_interface_display_sell_price, font=('Consolas', 10, 'bold')).grid(row=4, column=1, sticky='w', padx=5, pady=2)
        ttk.Label(results_lf, text="Total Sell Value:").grid(row=5, column=0, sticky='w', padx=5, pady=2)
        ttk.Label(results_lf, textvariable=self.ge_interface_display_sell_total, font=('Consolas', 10, 'bold')).grid(row=5, column=1, sticky='w', padx=5, pady=2)

    def _reset_search(self, *args):
        """Resets the search index whenever the search query is modified."""
        self.search_results = []
        self.current_search_index = -1

    def search_flowchart(self, event=None):
        """Finds and navigates to the next item matching the search query."""
        query = self.search_query.get().lower()
        if not query:
            return

        # If it's a new search, populate the results list
        if not self.search_results:
            # Search Steps
            for i, step in enumerate(self.steps):
                # Build a string of all searchable content for the step
                content = f"step {i+1} " + str(step.get('name', '')).lower()
                if step.get('path'):
                    content += " " + str(os.path.basename(step.get('path'))).lower()
                if step.get('key_to_press'):
                    content += " " + str(step.get('key_to_press')).lower()
                if step.get('text_to_type'):
                    content += " " + str(step.get('text_to_type')).lower()
                if step.get('ge_inject_name'):
                    content += " " + str(step.get('ge_inject_name')).lower()

                # Check for direct step number match or content match
                if query == str(i + 1) or query in content:
                    self.search_results.append({'type': 'step', 'index': i})
            
            # Search Annotations
            for i, note in enumerate(self.annotations):
                if query in str(note.get('text', '')).lower():
                    self.search_results.append({'type': 'note', 'index': i})

        if not self.search_results:
            self.log(f"Search: No results found for '{query}'.")
            return

        # Cycle to the next result
        self.current_search_index = (self.current_search_index + 1) % len(self.search_results)
        found_item = self.search_results[self.current_search_index]
        
        self.log(f"Search: Found item {self.current_search_index + 1}/{len(self.search_results)} at {found_item['type']} index {found_item['index']}.")

        # Highlight, update properties, and redraw
        self.selected_items = [found_item]
        self.populate_properties_panel()
        self.redraw_flowchart()
        
        # Scroll the canvas to the found item
        self.root.after(50, lambda: self._scroll_to_item(found_item))

    def _scroll_to_item(self, item):
        """Calculates position and scrolls the canvas to bring the item into view."""
        item_data = None
        if item['type'] == 'step':
            item_data = self.steps[item['index']]
        elif item['type'] == 'note':
            item_data = self.annotations[item['index']]

        if not item_data:
            return
        
        # Ensure canvas dimensions are up-to-date
        self.canvas.update_idletasks()

        item_x, item_y = item_data.get('x', 0) * self.zoom_factor, item_data.get('y', 0) * self.zoom_factor
        canvas_width = self.canvas.winfo_width()
        canvas_height = self.canvas.winfo_height()
        
        scroll_region = self.canvas.cget("scrollregion")
        if not scroll_region:
             return
        
        sr_parts = list(map(float, scroll_region.split()))
        sr_x, sr_y, sr_w, sr_h = sr_parts[0], sr_parts[1], sr_parts[2], sr_parts[3]
        total_w, total_h = sr_w - sr_x, sr_h - sr_y

        if total_w > 0 and total_h > 0:
            # Calculate desired top-left corner of the viewport
            target_x = item_x - (canvas_width / 2)
            target_y = item_y - (canvas_height / 2)
            
            # Convert to relative position (0.0 to 1.0) for moveto
            rel_x = max(0.0, target_x / total_w)
            rel_y = max(0.0, target_y / total_h)
            
            self.canvas.xview_moveto(rel_x)
            self.canvas.yview_moveto(rel_y)

    def clear_search(self):
        """Clears the search query and results."""
        self.search_query.set("")
        self._reset_search()
        self.selected_items = []
        self.populate_properties_panel()
        self.redraw_flowchart()
        self.log("Search cleared.")

    def _update_mouse_mode_visibility(self, *args):
        """Shows or hides mouse movement setting frames based on the selected mode."""
        mode = self.mouse_move_mode.get()
        
        if hasattr(self, 'regular_speed_frame'):
            if mode == 'Regular':
                self.regular_speed_frame.grid()
            else:
                self.regular_speed_frame.grid_remove()
                
        if hasattr(self, 'dynamic_speed_frame'):
            if mode == 'Dynamic':
                self.dynamic_speed_frame.grid()
            else:
                self.dynamic_speed_frame.grid_remove()

        if hasattr(self, 'pps_speed_frame'):
            if mode == 'Pixels Per Second':
                self.pps_speed_frame.grid()
            else:
                self.pps_speed_frame.grid_remove()

    def apply_global_settings(self):
        try:
            # Settings controlled by Radiobuttons, Checkbuttons, or Scales are updated
            # directly via their own variable bindings and do not need to be "applied"
            # by this function. We must skip them to avoid errors.
            keys_to_skip = ['mouse_move_mode', 'grid_visible', 'grid_latching', 'grid_opacity']

            for key, ui_var in self.global_settings_ui_vars.items():
                if key in keys_to_skip:
                    continue  # Skip settings that are not from Entry widgets.
                
                setting_map = self.global_settings_map[key]
                value = setting_map['type'](ui_var.get())
                setting_map['model'].set(value)
                
            self.log("Global settings applied successfully.", "green")
        except ValueError as e:
            messagebox.showerror("Invalid Input", f"Please check your global settings values.\nOne of the values is not a valid number.\n\nDetails: {e}")
            self.log(f"Failed to apply global settings: {e}", "red")

    def _sync_global_settings_ui_from_model(self):
        for key, ui_var in self.global_settings_ui_vars.items():
            ui_var.set(str(self.global_settings_map[key]['model'].get()))


