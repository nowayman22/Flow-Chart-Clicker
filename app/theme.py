import tkinter as tk
from tkinter import ttk
import ctypes

class ThemeMixin:
    def apply_theme(self):
        theme = {
            'bg': '#282C34', 'fg': '#ABB2BF', 'canvas': '#21252B', 'log': '#1E2228',
            'entry_bg': '#2C313A', 'entry_fg': '#D7D7D7', 'text': '#ABB2BF',
            'props_bg': '#2C313A', 'accent_green': '#98C379', 'accent_blue': '#61AFEF',
            'accent_grey': '#5C6370', 'accent_red': '#E06C75', 'accent_yellow': '#E5C07B',
            'accent_teal': '#56B6C2', 'accent_purple': '#C678DD', 'node_border': '#ABB2BF',
            'node_text_grey': '#8A92A0', 'btn_fg': '#FFFFFF', 'status_blue': '#61AFEF',
            'status_green': '#98C379', 'status_orange': '#D19A66', 'status_red': '#BE5046',
            'select_bg': '#3A4048', 'tree_heading_bg': '#2C313A', 'tree_heading_fg': '#ABB2BF', 'blk': '#b0192d'
        }
        self.current_theme = theme
        style = ttk.Style(); style.theme_use('clam')

        # General widget styling
        style.configure('.', background=theme['bg'], foreground=theme['fg'], fieldbackground=theme['entry_bg'], bordercolor=theme['node_border'], lightcolor=theme['bg'], darkcolor=theme['bg'])
        style.map('.', background=[('active', theme['entry_bg'])])

        # Specific widget configurations
        style.configure('TFrame', background=theme['bg'])
        style.configure('TLabel', background=theme['bg'], foreground=theme['fg'])
        style.configure('TLabelframe', background=theme['bg'], foreground=theme['fg'], bordercolor=theme['fg'])
        style.configure('TLabelframe.Label', background=theme['bg'], foreground=theme['fg'])
        style.configure('TEntry', fieldbackground=theme['entry_bg'], foreground=theme['entry_fg'], insertcolor=theme['text'])
        style.configure('TNotebook', background=theme['bg'], borderwidth=1)
        style.configure('TNotebook.Tab', background=theme['accent_grey'], foreground=theme['btn_fg'], padding=[8, 4])
        style.map('TNotebook.Tab', background=[('selected', theme['canvas']), ('active', theme['entry_bg'])], foreground=[('selected', theme['fg']), ('active', theme['fg'])])
        style.configure('TCheckbutton', background=theme['bg'], foreground=theme['fg'], indicatorcolor=theme['entry_bg'])
        style.map('TCheckbutton', 
                  indicatorcolor=[('selected', theme['status_green']), ('!selected', theme['entry_bg'])],
                  foreground=[('selected', theme['status_green'])],
                  background=[('active', theme['bg'])])
        style.configure('TRadiobutton', background=theme['bg'], foreground=theme['fg'])
        style.map('TRadiobutton', background=[('active', theme['entry_bg'])])
        style.configure('TCombobox', fieldbackground=theme['entry_bg'], foreground=theme['entry_fg'], arrowcolor=theme['fg'], background=theme['bg'])
        style.map('TCombobox', fieldbackground=[('readonly', theme['entry_bg'])])
        
        # Treeview styling for Deal Finder
        style.configure("Treeview", background=theme['entry_bg'], foreground=theme['entry_fg'], fieldbackground=theme['entry_bg'], rowheight=25)
        style.map("Treeview", background=[('selected', theme['select_bg'])], foreground=[('selected', theme['fg'])])
        style.configure("Treeview.Heading", background=theme['tree_heading_bg'], foreground=theme['tree_heading_fg'], font=('Helvetica', 10, 'bold'))
        style.map("Treeview.Heading", background=[('active', theme['accent_grey'])])

        self.root.config(bg=theme['bg'])
        self.update_widget_colors_recursive(self.root, theme)
        self.populate_properties_panel()
        self.redraw_flowchart()
        state_color_map = { 'blue': theme['status_blue'], 'green': theme['status_green'], 'orange': theme['status_orange'], 'red': theme['status_red'] }
        new_color = state_color_map.get(self.status_label_color_state, theme['status_blue'])
        self.status_label.config(foreground=new_color)
        
        self.root.after(10, self._set_title_bar_color)

    def _set_title_bar_color(self):
        """
        Applies the custom title bar color after a short delay to ensure the window is ready.
        This is a Windows-only feature.
        """
        try:
            hwnd = ctypes.windll.user32.GetParent(self.root.winfo_id())
            # First, set the window to dark mode (attribute 20)
            true_value = ctypes.c_bool(True)
            ctypes.windll.dwmapi.DwmSetWindowAttribute(hwnd, 20, ctypes.byref(true_value), ctypes.sizeof(true_value))
            
            # Then, set the custom colors for caption (35) and border (34)
            title_bar_color = 0x003A312C # Dark Grey
            ctypes.windll.dwmapi.DwmSetWindowAttribute(hwnd, 35, ctypes.byref(ctypes.c_int(title_bar_color)), ctypes.sizeof(ctypes.c_int(title_bar_color)))
            ctypes.windll.dwmapi.DwmSetWindowAttribute(hwnd, 34, ctypes.byref(ctypes.c_int(title_bar_color)), ctypes.sizeof(ctypes.c_int(title_bar_color)))
        except (AttributeError, TypeError):
            # Fails gracefully on non-Windows systems or if the call fails
            pass

    def update_widget_colors_recursive(self, widget, theme):
        widget_class = widget.winfo_class()
        try:
            if widget_class not in ('Canvas', 'Entry', 'Toplevel', 'ScrolledText', 'Text', 'TMenubutton', 'Treeview'):
                widget.config(bg=theme['bg'])
            if 'fg' in widget.config() and widget_class not in ('Button'):
                widget.config(fg=theme['fg'])
            
            if widget_class in ('Entry', 'Text'):
                widget.config(bg=theme['entry_bg'], fg=theme['entry_fg'], insertbackground=theme['text'])
            elif widget_class == 'ScrolledText':
                widget.config(bg=theme['log'], fg=theme['text'], insertbackground=theme['text'])
            elif widget_class == 'Radiobutton':
                widget.config(selectcolor=theme['entry_bg'], activebackground=theme['entry_bg'], highlightbackground=theme['bg'], highlightcolor=theme['bg'])
            elif widget_class == 'TMenubutton':
                widget.config(background=theme['entry_bg'])
            elif widget_class == 'Button':
                color_map = {
                    "Color Step": 'accent_green', "PNG Step": 'accent_blue', "Click / Press Step": 'accent_grey',
                    "Logical Step": 'accent_yellow', "Reset All": 'accent_red',
                    "Delete": 'accent_red', "Duplicate": 'accent_grey', "Snip": 'accent_yellow',
                    "Apply Changes": 'accent_green', "Apply Global Settings": 'accent_green',
                    "Run Test": 'accent_green', "Add Note": 'accent_teal', "Find Profitable": 'accent_teal',
                    "Fetch & Calculate Price": 'accent_purple'
                }
                bg_key = next((key for name, key in color_map.items() if name in widget.cget('text')), 'accent_grey')
                widget.config(bg=theme[bg_key], fg=theme['btn_fg'], activebackground=theme['entry_bg'], activeforeground=theme['fg'], relief=tk.FLAT, borderwidth=0)
        except (tk.TclError, AttributeError):
            pass # Ignore errors for widgets that don't support these properties
        for child in widget.winfo_children():
            self.update_widget_colors_recursive(child, theme)

