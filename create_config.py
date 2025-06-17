import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import json
import os

# --- Default Templates based on the user's provided JF_config.json structure ---

INVOICE_TEMPLATE = {
    "start_row": 21,
    "add_blank_before_footer": False,
    "static_content_before_footer": {
        "2": "HS CODE: 4202.92.00"
    },
    "header_to_write": [
        {"row": 0, "col": 0, "text": "Mark & Nº", "id": "col_static", "rowspan": 1, "colspan": 1},
        {"row": 0, "col": 1, "text": "P.O. Nº", "id": "col_po", "rowspan": 1, "colspan": 1},
        {"row": 0, "col": 2, "text": "ITEM Nº", "id": "col_item", "rowspan": 1, "colspan": 1},
        {"row": 0, "col": 3, "text": "Description", "id": "col_desc", "rowspan": 1, "colspan": 1},
        {"row": 0, "col": 4, "text": "Quantity", "id": "col_qty_sf", "rowspan": 1, "colspan": 1},
        {"row": 0, "col": 5, "text": "Unit price (USD)", "id": "col_unit_price", "rowspan": 1, "colspan": 1},
        {"row": 0, "col": 6, "text": "Amount (USD)", "id": "col_amount", "rowspan": 1, "colspan": 1}
    ],
    "mappings": {
        "po": {"key_index": 0, "id": "col_po"},
        "item": {"key_index": 1, "id": "col_item"},
        "description": {"key_index": 3, "id": "col_desc", "fallback_on_none": "LEATHER"},
        "sqft": {"value_key": "sqft_sum", "id": "col_qty_sf"},
        "unit_price": {"key_index": 2, "id": "col_unit_price"},
        "amount": {
            "id": "col_amount", "type": "formula",
            "formula_template": "{col_ref_1}{row} * {col_ref_0}{row}",
            "inputs": ["col_qty_sf", "col_unit_price"]
        },
        "initial_static": {
            "type": "initial_static_rows", "column_header_id": "col_static",
            "values": ["VENDOR#:", "Des: LEATHER", "MADE IN CAMBODIA"]
        }
    },
    "data_cell_merging_rule": {"col_item": {"rowspan": 1}},
    "weight_summary_config": {"enabled": True, "label_col_id": "col_po", "value_col_id": "col_item"},
    "footer_configurations": {
        "total_text": "TOTAL:", "total_text_column_id": "col_po", "pallet_count_column_id": "col_desc",
        "sum_column_ids": ["col_qty_pcs", "col_qty_sf", "col_gross", "col_net", "col_cbm", "col_amount"],
        "number_formats": {
            "col_qty_pcs": {"number_format": "#,##0"}, "col_qty_sf": {"number_format": "#,##0.00"},
            "col_gross": {"number_format": "#,##0.00"}, "col_net": {"number_format": "#,##0.00"},
            "col_cbm": {"number_format": "0.00"}, "col_amount": {"number_format": "#,##0.00"}
        },
        "style": {
            "font": {"name": "Times New Roman", "size": 12, "bold": True},
            "alignment": {"horizontal": "center", "vertical": "center"}, "border": {"apply": True}
        },
        "merge_rules": [{"start_column_id": "col_po", "colspan": 1}]
    },
    "styling": {
        "force_text_format_ids": ["col_po", "col_item"],
        "column_ids_with_full_grid": ["col_po", "col_desc", "col_item", "col_qty_sf", "col_unit_price", "col_amount"],
        "column_id_styles": {
            "col_unit_price": {"number_format": "#,##0.00"}, "col_amount": {"number_format": "#,##0.00"},
            "col_qty_sf": {"number_format": "#,##0.00"}, "col_desc": {"alignment": {"horizontal": "center"}}
        },
        "column_id_widths": {"col_po": 28, "col_desc": 20},
        "default_font": {"name": "Times New Roman", "size": 12},
        "header_font": {"name": "Times New Roman", "size": 12, "bold": True},
        "default_alignment": {"horizontal": "center", "vertical": "center", "wrap_text": True},
        "header_alignment": {"horizontal": "center", "vertical": "center", "wrap_text": True},
        "row_heights": {"header": 35, "data_default": 35, "footer": 35}
    }
}

CONTRACT_TEMPLATE = {
    "start_row": 15,
    "header_to_write": [
        {"row": 0, "col": 0, "text": "No.", "id": "col_no", "rowspan": 1, "colspan": 1},
        {"row": 0, "col": 1, "text": "ITEM Nº", "id": "col_item", "rowspan": 1, "colspan": 1},
        {"row": 0, "col": 2, "text": "Quantity", "id": "col_qty_sf", "rowspan": 1, "colspan": 1},
        {"row": 0, "col": 3, "text": "Unit Price(USD)", "id": "col_unit_price", "rowspan": 1, "colspan": 1},
        {"row": 0, "col": 4, "text": "Total value(USD)", "id": "col_amount", "rowspan": 1, "colspan": 1}
    ],
    "mappings": {
        "po": {"key_index": 0, "id": "col_po"}, "item": {"key_index": 1, "id": "col_item"},
        "desc": {"key_index": 3, "id": "col_desc", "fallback_on_none": "LEATHER"},
        "sqft": {"value_key": "sqft_sum", "id": "col_qty_sf"},
        "unit_price": {"key_index": 2, "id": "col_unit_price"},
        "amount": {
            "id": "col_amount", "type": "formula",
            "formula_template": "{col_ref_1}{row} * {col_ref_0}{row}",
            "inputs": ["col_qty_sf", "col_unit_price"]
        }
    },
    "footer_configurations": {
        "total_text": "TOTAL:", "total_text_column_id": "col_no",
        "sum_column_ids": ["col_qty_pcs", "col_qty_sf", "col_gross", "col_net", "col_cbm", "col_amount"],
        "style": {
            "font": {"name": "Times New Roman", "size": 12, "bold": True},
            "alignment": {"horizontal": "center", "vertical": "center"}, "border": {"apply": True}
        },
        "number_formats": {
            "col_qty_pcs": {"number_format": "#,##0"}, "col_qty_sf": {"number_format": "#,##0.00"},
            "col_gross": {"number_format": "#,##0.00"}, "col_net": {"number_format": "#,##0.00"},
            "col_cbm": {"number_format": "0.00"}, "col_amount": {"number_format": "#,##0.00"}
        },
        "merge_rules": [{"start_column_id": "col_no", "colspan": 2}]
    },
    "styling": {
        "header_pattern_fill": {"fill_type": "solid", "start_color": "D3D3D3"},
        "force_text_format_ids": ["col_po", "col_item", "col_no"],
        "column_ids_with_full_grid": ["col_no", "col_po", "col_item", "col_desc", "col_qty_sf", "col_unit_price", "col_amount"],
        "default_font": {"name": "Times New Roman", "size": 14},
        "header_font": {"name": "Times New Roman", "size": 16, "bold": True},
        "default_alignment": {"horizontal": "center", "vertical": "center", "wrap_text": True},
        "header_alignment": {"horizontal": "center", "vertical": "center", "wrap_text": True},
        "column_id_styles": {
            "col_amount": {"number_format": "#,##0.00"}, "col_unit_price": {"number_format": "#,##0.00"},
            "col_qty_sf": {"number_format": "#,##0.00"}, "col_desc": {"alignment": {"horizontal": "left"}}
        },
        "column_id_widths": {"col_no": 14, "col_desc": 35, "col_qty_sf": 27, "col_unit_price": 28, "col_amount": 47},
        "row_heights": {"header": 36, "data_default": 30}
    }
}

PACKING_LIST_TEMPLATE = {
    "start_row": 21,
    "add_blank_before_footer": True,
    "static_content_before_footer": {"2": "LEATHER (HS.CODE: 4107.12.00)"},
    "summary": True,
    "merge_rules_before_footer": {"2": 2},
    "header_to_write": [
        {"row": 0, "col": 0, "text": "Mark & Nº", "id": "col_static", "rowspan": 2, "colspan": 1},
        {"row": 0, "col": 1, "text": "P.O Nº", "id": "col_po", "rowspan": 2, "colspan": 1},
        {"row": 0, "col": 2, "text": "ITEM Nº", "id": "col_item", "rowspan": 2, "colspan": 1},
        {"row": 0, "col": 3, "text": "Description", "id": "col_desc", "rowspan": 2, "colspan": 1},
        {"row": 0, "col": 4, "text": "Quantity", "rowspan": 1, "colspan": 2},
        {"row": 0, "col": 6, "text": "G.W (kgs)", "id": "col_gross", "rowspan": 2, "colspan": 1},
        {"row": 0, "col": 7, "text": "N.W (kgs)", "id": "col_net", "rowspan": 2, "colspan": 1},
        {"row": 0, "col": 8, "text": "CBM", "id": "col_cbm", "rowspan": 2, "colspan": 1},
        {"row": 1, "col": 4, "text": "PCS", "id": "col_qty_pcs", "rowspan": 1, "colspan": 1},
        {"row": 1, "col": 5, "text": "SF", "id": "col_qty_sf", "rowspan": 1, "colspan": 1}
    ],
    "mappings": {
        "initial_static": {
            "type": "initial_static_rows", "column_header_id": "col_static",
            "values": ["VENDOR#:", "Des: LEATHER", "Case Qty:", "MADE IN CAMBODIA"]
        },
        "data_map": {
            "po": {"id": "col_po"}, "item": {"id": "col_item"},
            "description": {"id": "col_desc", "fallback_on_none": "LEATHER"},
            "pcs": {"id": "col_qty_pcs"}, "sqft": {"id": "col_qty_sf"},
            "net": {"id": "col_net"}, "gross": {"id": "col_gross"}, "cbm": {"id": "col_cbm"}
        }
    },
    "footer_configurations": {
        "total_text": "TOTAL:", "total_text_column_id": "col_po", "pallet_count_column_id": "col_desc",
        "sum_column_ids": ["col_qty_pcs", "col_qty_sf", "col_gross", "col_net", "col_cbm"],
        "number_formats": {
            "col_qty_pcs": {"number_format": "#,##0"}, "col_qty_sf": {"number_format": "#,##0.00"},
            "col_gross": {"number_format": "#,##0.00"}, "col_net": {"number_format": "#,##0.00"},
            "col_cbm": {"number_format": "0.00"}, "col_amount": {"number_format": "#,##0.00"}
        },
        "style": {
            "font": {"name": "Times New Roman", "size": 12, "bold": True},
            "alignment": {"horizontal": "center", "vertical": "center"}, "border": {"apply": True}
        },
        "merge_rules": [{"start_column_id": "col_po", "colspan": 1}]
    },
    "styling": {
        "header_pattern_fill": {"fill_type": "solid", "start_color": "D3D3D3"},
        "force_text_format_ids": ["col_po", "col_item", "col_desc"],
        "column_ids_with_full_grid": ["col_po", "col_item", "col_desc", "col_qty_pcs", "col_qty_sf", "col_net", "col_gross", "col_cbm"],
        "default_font": {"name": "Times New Roman", "size": 12},
        "default_alignment": {"horizontal": "center", "vertical": "center", "wrap_text": True},
        "header_font": {"name": "Times New Roman", "size": 12, "bold": True},
        "header_alignment": {"horizontal": "center", "vertical": "center", "wrap_text": True},
        "column_id_styles": {
            "col_static": {"alignment": {"horizontal": "left", "vertical": "top"}, "font": {"size": 12}},
            "col_desc": {"alignment": {"horizontal": "center"}},
            "col_net": {"number_format": "#,##0.00"}, "col_gross": {"number_format": "#,##0.00"},
            "col_cbm": {"number_format": "0.00"}, "col_qty_pcs": {"number_format": "#,##0"},
            "col_qty_sf": {"number_format": "#,##0.00"}
        },
        "column_id_widths": {
            "col_static": 24.71, "col_po": 17, "col_item": 22.14, "col_desc": 26, "col_qty_pcs": 15,
            "col_qty_sf": 15, "col_net": 15, "col_gross": 15, "col_cbm": 15
        },
        "row_heights": {"header": 27, "data_default": 27, "before_footer": 27}
    }
}

DEFAULT_CONFIG_TEMPLATE = {
    "sheets_to_process": ["Invoice", "Contract", "Packing list"],
    "sheet_data_map": {
        "Invoice": "aggregation",
        "Contract": "aggregation",
        "Packing list": "processed_tables_multi"
    },
    "data_mapping": {
        "Invoice": INVOICE_TEMPLATE,
        "Contract": CONTRACT_TEMPLATE,
        "Packing list": PACKING_LIST_TEMPLATE
    }
}


# --- Helper Classes ---

class ToolTip:
    """Creates a tooltip for a given widget."""
    def __init__(self, widget, text):
        self.widget = widget
        self.text = text
        self.tooltip_window = None
        self.widget.bind("<Enter>", self.show_tooltip)
        self.widget.bind("<Leave>", self.hide_tooltip)

    def show_tooltip(self, event):
        x, y, _, _ = self.widget.bbox("insert")
        x += self.widget.winfo_rootx() + 25
        y += self.widget.winfo_rooty() + 25
        self.tooltip_window = tk.Toplevel(self.widget)
        self.tooltip_window.wm_overrideredirect(True)
        self.tooltip_window.wm_geometry(f"+{x}+{y}")
        label = tk.Label(self.tooltip_window, text=self.text, justify='left',
                         background="#ffffe0", relief='solid', borderwidth=1,
                         font=("tahoma", "8", "normal"))
        label.pack(ipadx=1)

    def hide_tooltip(self, event):
        if self.tooltip_window:
            self.tooltip_window.destroy()
        self.tooltip_window = None

class DictionaryEditor(tk.Toplevel):
    """A Toplevel window for editing key-value pairs of a dictionary."""
    def __init__(self, parent, title, data_dict, on_save):
        super().__init__(parent)
        self.title(title)
        self.geometry("500x400")
        self.data_dict = data_dict.copy()
        self.on_save = on_save
        self.widgets = []

        canvas = tk.Canvas(self)
        scrollbar = ttk.Scrollbar(self, orient="vertical", command=canvas.yview)
        self.scrollable_frame = ttk.Frame(canvas)
        self.scrollable_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        
        tk.Label(self.scrollable_frame, text="Key", font=("Helvetica", 10, "bold")).grid(row=0, column=0, padx=5)
        tk.Label(self.scrollable_frame, text="Value", font=("Helvetica", 10, "bold")).grid(row=0, column=1, padx=5)
        
        self.populate_rows()
        
        button_frame = tk.Frame(self)
        tk.Button(button_frame, text="+ Add", command=self.add_row).pack(side="left", padx=5)
        tk.Button(button_frame, text="Save & Close", command=self.save_and_close).pack(side="right", padx=5)
        
        canvas.pack(side="top", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        button_frame.pack(side="bottom", fill="x", pady=5)

        self.transient(parent)
        self.grab_set()

    def populate_rows(self):
        for widget_row in self.widgets:
            widget_row["key"].destroy()
            widget_row["value"].destroy()
            widget_row["remove_btn"].destroy()
        self.widgets.clear()
        
        for i, (key, value) in enumerate(self.data_dict.items()):
            key_entry = tk.Entry(self.scrollable_frame)
            key_entry.insert(0, key)
            key_entry.grid(row=i + 1, column=0, padx=5, pady=2, sticky="ew")

            value_entry = tk.Entry(self.scrollable_frame)
            try:
                value_str = json.dumps(value)
            except TypeError:
                value_str = str(value) 
            value_entry.insert(0, value_str)
            value_entry.grid(row=i + 1, column=1, padx=5, pady=2, sticky="ew")

            remove_btn = tk.Button(self.scrollable_frame, text="X", fg="red", command=lambda k=key: self.remove_row(k))
            remove_btn.grid(row=i + 1, column=2, padx=5)
            
            self.widgets.append({"key": key_entry, "value": value_entry, "remove_btn": remove_btn})
        self.scrollable_frame.grid_columnconfigure(1, weight=1)

    def add_row(self):
        self.data_dict[f"new_key_{len(self.data_dict)}"] = "new_value"
        self.populate_rows()

    def remove_row(self, key_to_remove):
        if key_to_remove in self.data_dict:
            del self.data_dict[key_to_remove]
        self.populate_rows()
        
    def save_and_close(self):
        new_dict = {}
        for row in self.widgets:
            key = row["key"].get().strip()
            value_str = row["value"].get().strip()
            if key:
                try:
                    value = json.loads(value_str)
                except json.JSONDecodeError:
                    value = value_str
                new_dict[key] = value
        self.on_save(new_dict)
        self.destroy()

class JsonConfigEditorApp:
    """The main application class for the Invoice Configuration Editor."""
    def __init__(self, root_window):
        self.root = root_window
        self.root.title("Invoice Configuration Editor")
        self.root.geometry("1400x850")

        self.config_data = {}
        self.current_filepath = None
        self.selected_sheet_name = tk.StringVar()
        self.temp_dict_data = {}
        
        self.sheet_start_row = None
        self.header_widgets = []
        self.mapping_widgets = []

        self._setup_menu()
        self._setup_main_layout()
        self.new_config()

    def _setup_menu(self):
        self.menubar = tk.Menu(self.root)
        self.root.config(menu=self.menubar)
        file_menu = tk.Menu(self.menubar, tearoff=0)
        file_menu.add_command(label="New Config", command=self.new_config)
        file_menu.add_command(label="Open...", command=self.open_config)
        file_menu.add_command(label="Save", command=self.save_config)
        file_menu.add_command(label="Save As...", command=self.save_config_as)
        file_menu.add_separator()
        file_menu.add_command(label="Exit", command=self.root.quit)
        self.menubar.add_cascade(label="File", menu=file_menu)

    def _setup_main_layout(self):
        left_panel = tk.Frame(self.root, width=250, relief=tk.RIDGE, bd=2)
        left_panel.pack(side=tk.LEFT, fill=tk.Y, padx=5, pady=5)
        left_panel.pack_propagate(False)
        tk.Label(left_panel, text="Select Sheet to Edit", font=("Helvetica", 12, "bold")).pack(pady=10)
        
        self.sheet_listbox = tk.Listbox(left_panel, exportselection=False, height=3)
        self.sheet_listbox.pack(expand=False, fill=tk.X, padx=5, pady=5)
        self.sheet_listbox.bind("<<ListboxSelect>>", self.on_sheet_select)

        right_panel = tk.Frame(self.root)
        right_panel.pack(side=tk.RIGHT, expand=True, fill=tk.BOTH, padx=5, pady=5)
        self.notebook = ttk.Notebook(right_panel)
        self.notebook.pack(expand=True, fill=tk.BOTH)

    def new_config(self):
        if messagebox.askokcancel("New Config", "Create a new configuration? This will clear any current edits."):
            self.config_data = json.loads(json.dumps(DEFAULT_CONFIG_TEMPLATE))
            self.current_filepath = None
            self.root.title("Invoice Configuration Editor - New Config")
            self.update_sheet_listbox()
            self.clear_notebook_tabs()

    def open_config(self):
        filepath = filedialog.askopenfilename(title="Open Configuration File", filetypes=[("JSON files", "*.json"), ("All files", "*.*")])
        if not filepath: return
        try:
            with open(filepath, 'r', encoding='utf-8') as f: self.config_data = json.load(f)
            self.current_filepath = filepath
            self.root.title(f"Invoice Configuration Editor - {os.path.basename(filepath)}")
            
            # This will repopulate the sheet list on the left AND trigger on_sheet_select,
            # which correctly populates the UI for the first sheet.
            self.update_sheet_listbox()
            
            ### FIX: This line was incorrectly erasing the UI after it was drawn. It has been removed.
            # self.clear_notebook_tabs() 
            
        except Exception as e:
            messagebox.showerror("Error Opening File", f"Failed to open or parse file:\n{e}")

    def save_config(self):
        if not self.current_filepath: self.save_config_as()
        else: self._save_to_file(self.current_filepath)

    def save_config_as(self):
        filepath = filedialog.asksaveasfilename(title="Save Configuration As", defaultextension=".json", initialfile="JF_config.json", filetypes=[("JSON files", "*.json"), ("All files", "*.*")])
        if not filepath: return
        self._save_to_file(filepath)
        self.current_filepath = filepath
        self.root.title(f"Invoice Configuration Editor - {os.path.basename(filepath)}")

    def _save_to_file(self, filepath):
        try:
            self.update_config_from_ui()
            with open(filepath, 'w', encoding='utf-8') as f: json.dump(self.config_data, f, indent=4)
            messagebox.showinfo("Save Successful", f"Configuration saved to:\n{filepath}")
        except Exception as e: messagebox.showerror("Error Saving File", f"Failed to save file:\n{e}")

    def update_sheet_listbox(self):
        self.sheet_listbox.delete(0, tk.END)
        # Make sure there is data to process
        if not self.config_data: return
        for sheet_name in self.config_data.get("sheets_to_process", []):
            self.sheet_listbox.insert(tk.END, sheet_name)
        if self.sheet_listbox.size() > 0:
            self.sheet_listbox.selection_set(0)
            self.on_sheet_select(None)

    def on_sheet_select(self, event):
        selected_indices = self.sheet_listbox.curselection()
        if not selected_indices: return
        
        if self.selected_sheet_name.get(): 
            self.update_config_from_ui()
            
        sheet_name = self.sheet_listbox.get(selected_indices[0])
        self.selected_sheet_name.set(sheet_name)
        self.populate_notebook_for_sheet(sheet_name)

    def clear_notebook_tabs(self):
        for i in self.notebook.tabs(): self.notebook.forget(i)
        self.temp_dict_data.clear()

    def populate_notebook_for_sheet(self, sheet_name):
        self.clear_notebook_tabs()
        sheet_conf = self.config_data.get("data_mapping", {}).get(sheet_name)
        if not sheet_conf: 
            messagebox.showwarning("Config Error", f"No configuration found for sheet '{sheet_name}' in 'data_mapping'.")
            return
        
        self.temp_dict_data['column_id_styles'] = sheet_conf.get('styling', {}).get('column_id_styles', {})
        self.temp_dict_data['column_id_widths'] = sheet_conf.get('styling', {}).get('column_id_widths', {})
        self.temp_dict_data['row_heights'] = sheet_conf.get('styling', {}).get('row_heights', {})
        self.temp_dict_data['number_formats'] = sheet_conf.get('footer_configurations', {}).get('number_formats', {})
        self.temp_dict_data['static_content_before_footer'] = sheet_conf.get('static_content_before_footer', {})
        self.temp_dict_data['data_cell_merging_rule'] = sheet_conf.get('data_cell_merging_rule', {})
        self.temp_dict_data['merge_rules_before_footer'] = sheet_conf.get('merge_rules_before_footer', {})

        self.create_sheet_settings_tab(sheet_conf)
        self.create_header_tab(sheet_conf)
        self.create_mappings_tab(sheet_conf)
        self.create_styling_tab(sheet_conf)
        self.create_footer_tab(sheet_conf)
        self.create_advanced_tab(sheet_conf)
    
    def create_sheet_settings_tab(self, sheet_conf):
        tab = ttk.Frame(self.notebook, padding=10)
        self.notebook.add(tab, text="Sheet Settings")
        
        tk.Label(tab, text="Start Row:").grid(row=0, column=0, sticky="w", padx=5, pady=5)
        self.sheet_start_row = tk.Entry(tab)
        self.sheet_start_row.insert(0, sheet_conf.get("start_row", ""))
        self.sheet_start_row.grid(row=0, column=1, sticky="ew", padx=5, pady=5)
        ToolTip(self.sheet_start_row, "The row number where the script should start writing the header.")

        tk.Label(tab, text="Data Source Type:").grid(row=1, column=0, sticky="w", padx=5, pady=5)
        self.sheet_data_source = tk.Entry(tab)
        self.sheet_data_source.insert(0, self.config_data.get("sheet_data_map", {}).get(self.selected_sheet_name.get(), ""))
        self.sheet_data_source.grid(row=1, column=1, sticky="ew", padx=5, pady=5)
        ToolTip(self.sheet_data_source, "The key from 'sheet_data_map' that defines data source type.")
        
        self.sheet_summary_var = tk.BooleanVar(value=sheet_conf.get("summary", False))
        tk.Checkbutton(tab, text="Enable Summary Rows", variable=self.sheet_summary_var).grid(row=2, column=0, columnspan=2, sticky="w", pady=5)
        
        self.sheet_bbf_var = tk.BooleanVar(value=sheet_conf.get("add_blank_before_footer", False))
        tk.Checkbutton(tab, text="Add Blank Row Before Footer", variable=self.sheet_bbf_var).grid(row=3, column=0, columnspan=2, sticky="w", pady=5)

    def create_header_tab(self, sheet_conf):
        tab = ttk.Frame(self.notebook, padding=10)
        self.notebook.add(tab, text="Header")
        canvas = tk.Canvas(tab)
        scrollbar = ttk.Scrollbar(tab, orient="vertical", command=canvas.yview)
        frame = ttk.Frame(canvas)
        frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        self.header_widgets = []
        headers = ["Row", "Col", "Text", "ID", "Row Span", "Col Span", ""]
        for i, h in enumerate(headers): tk.Label(frame, text=h, font=("Helvetica", 10, "bold")).grid(row=0, column=i, padx=5)
        for i, item in enumerate(sheet_conf.get("header_to_write", [])): self.add_header_row_ui(frame, i + 1, item)
        tk.Button(frame, text="+ Add Header Item", command=lambda: self.add_header_row_ui(frame, len(self.header_widgets) + 1, {})).grid(row=len(self.header_widgets) + 2, column=0, columnspan=len(headers), pady=10)

    def add_header_row_ui(self, parent, row_num, item_data):
        widgets = {}
        e_row = tk.Entry(parent, width=5); e_row.insert(0, item_data.get("row", "0")); e_row.grid(row=row_num, column=0); widgets["row"] = e_row
        e_col = tk.Entry(parent, width=5); e_col.insert(0, item_data.get("col", "0")); e_col.grid(row=row_num, column=1); widgets["col"] = e_col
        e_text = tk.Entry(parent, width=30); e_text.insert(0, item_data.get("text", "")); e_text.grid(row=row_num, column=2); widgets["text"] = e_text
        e_id = tk.Entry(parent, width=20); e_id.insert(0, item_data.get("id", "")); e_id.grid(row=row_num, column=3); widgets["id"] = e_id
        e_rs = tk.Entry(parent, width=8); e_rs.insert(0, item_data.get("rowspan", "1")); e_rs.grid(row=row_num, column=4); widgets["rowspan"] = e_rs
        e_cs = tk.Entry(parent, width=8); e_cs.insert(0, item_data.get("colspan", "1")); e_cs.grid(row=row_num, column=5); widgets["colspan"] = e_cs
        remove_button = tk.Button(parent, text="X", fg="red", command=lambda w=widgets: self.remove_row_ui(w, self.header_widgets))
        remove_button.grid(row=row_num, column=6)
        widgets["remove_btn"] = remove_button
        self.header_widgets.append(widgets)

    def remove_row_ui(self, row_to_remove, widget_list):
        container = row_to_remove.get("_row_frame")
        if isinstance(container, tk.Widget):
            container.destroy()
        else:
            for item in row_to_remove.values():
                if isinstance(item, tk.Widget):
                    item.destroy()
        
        if row_to_remove in widget_list:
            widget_list.remove(row_to_remove)
        
    def create_mappings_tab(self, sheet_conf):
        tab = ttk.Frame(self.notebook, padding=10)
        self.notebook.add(tab, text="Mappings")
        canvas = tk.Canvas(tab)
        scrollbar = ttk.Scrollbar(tab, orient="vertical", command=canvas.yview)
        self.mappings_scrollable_frame = ttk.Frame(canvas)
        self.mappings_scrollable_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=self.mappings_scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        self.mapping_widgets = []
        headers = ["Mapping Name", "Target Col ID", "Mapping Type", "Rule Details", "Fallback", ""]
        for i, h in enumerate(headers): tk.Label(self.mappings_scrollable_frame, text=h, font=("Helvetica", 10, "bold")).grid(row=0, column=i, padx=5)
        mappings = sheet_conf.get("mappings", {})
        for i, (name, rule) in enumerate(mappings.items()): self.add_mapping_row_ui(i + 1, name, rule)
        tk.Button(self.mappings_scrollable_frame, text="+ Add Mapping Rule", command=self.add_new_mapping_rule).grid(row=len(self.mapping_widgets) + 2, column=0, columnspan=len(headers), pady=10)

    def add_mapping_row_ui(self, row_num, name, rule):
        parent = self.mappings_scrollable_frame
        row_frame = tk.Frame(parent)
        row_frame.grid(row=row_num, column=0, columnspan=6, sticky='ew', pady=2)
        widgets = {"_row_frame": row_frame, "_rule_data": rule}
        
        name_entry = tk.Entry(row_frame, width=15); name_entry.insert(0, name); name_entry.pack(side="left", padx=2); widgets["name"] = name_entry
        id_entry = tk.Entry(row_frame, width=15); id_entry.insert(0, rule.get("id", "")); id_entry.pack(side="left", padx=2); widgets["id"] = id_entry
        
        rule_type = self._determine_rule_type(rule)
        type_combo = ttk.Combobox(row_frame, values=["From Data Key", "From Data Value", "Formula", "Static Rows", "Data Map"], width=15, state="readonly"); type_combo.set(rule_type); type_combo.pack(side="left", padx=2); widgets["type_combo"] = type_combo
        
        details_frame = tk.Frame(row_frame); details_frame.pack(side="left", padx=2, fill="x", expand=True); widgets["_details_frame"] = details_frame
        fallback_entry = tk.Entry(row_frame, width=15); fallback_entry.insert(0, rule.get("fallback_on_none", "")); fallback_entry.pack(side="left", padx=2); widgets["fallback"] = fallback_entry
        
        remove_button = tk.Button(row_frame, text="X", fg="red", command=lambda w=widgets: self.remove_row_ui(w, self.mapping_widgets))
        remove_button.pack(side="left", padx=5)
        
        type_combo.bind("<<ComboboxSelected>>", lambda event, w=widgets: self._update_mapping_rule_details_ui(w))
        self.mapping_widgets.append(widgets)
        self._update_mapping_rule_details_ui(widgets)

    def _determine_rule_type(self, rule):
        if rule.get("type") == "formula": return "Formula"
        if rule.get("type") == "initial_static_rows": return "Static Rows"
        if "data_map" in rule: return "Data Map" 
        if "key_index" in rule: return "From Data Key"
        if "value_key" in rule: return "From Data Value"
        return "From Data Value"

    def _update_mapping_rule_details_ui(self, widget_dict):
        details_frame = widget_dict["_details_frame"]
        for child in details_frame.winfo_children(): child.destroy()
        
        keys_to_clear = ["key_index", "value_key", "formula_template", "inputs"]
        for key in keys_to_clear:
            widget_dict.pop(key, None)

        rule_type = widget_dict["type_combo"].get()
        rule = widget_dict.get("_rule_data", {})
        
        if rule_type == "From Data Key":
            tk.Label(details_frame, text="Key Index:").pack(side="left")
            entry = tk.Entry(details_frame, width=10); entry.insert(0, rule.get("key_index", "")); entry.pack(side="left"); widget_dict["key_index"] = entry
        elif rule_type == "From Data Value":
            tk.Label(details_frame, text="Value Key:").pack(side="left")
            entry = tk.Entry(details_frame, width=20); entry.insert(0, rule.get("value_key", "")); entry.pack(side="left"); widget_dict["value_key"] = entry
        elif rule_type == "Formula":
            tk.Label(details_frame, text="Template:").pack(side="left")
            template_entry = tk.Entry(details_frame, width=25); template_entry.insert(0, rule.get("formula_template", "")); template_entry.pack(side="left", padx=(0, 5)); widget_dict["formula_template"] = template_entry
            tk.Label(details_frame, text="Inputs:").pack(side="left")
            inputs_entry = tk.Entry(details_frame, width=25); inputs_entry.insert(0, ", ".join(rule.get("inputs", []))); inputs_entry.pack(side="left"); widget_dict["inputs"] = inputs_entry
            ToolTip(inputs_entry, "Comma-separated list of column IDs")
        elif rule_type in ["Static Rows", "Data Map"]:
            tk.Label(details_frame, text="Defined in Mappings").pack(side="left")

    def add_new_mapping_rule(self):
        new_row_num = len(self.mapping_widgets) + 1
        self.add_mapping_row_ui(new_row_num, f"new_mapping_{new_row_num}", {})

    def create_styling_tab(self, sheet_conf):
        tab = ttk.Frame(self.notebook, padding=10)
        self.notebook.add(tab, text="Styling")
        styling_conf = sheet_conf.get("styling", {})
        
        tk.Label(tab, text="Force Text Format IDs:").grid(row=0, column=0, sticky="w", pady=2)
        self.style_force_text = tk.Entry(tab, width=80); self.style_force_text.insert(0, ", ".join(styling_conf.get("force_text_format_ids", []))); self.style_force_text.grid(row=0, column=1, pady=2, sticky="ew")
        
        tk.Label(tab, text="Column IDs with Full Grid:").grid(row=1, column=0, sticky="w", pady=2)
        self.style_full_grid = tk.Entry(tab, width=80); self.style_full_grid.insert(0, ", ".join(styling_conf.get("column_ids_with_full_grid", []))); self.style_full_grid.grid(row=1, column=1, pady=2, sticky="ew")

        tk.Label(tab, text="Column ID Styles:").grid(row=2, column=0, sticky="w", pady=2); tk.Button(tab, text="Edit...", command=lambda: self.open_dict_editor("Column ID Styles", "column_id_styles")).grid(row=2, column=1, sticky="w")
        tk.Label(tab, text="Column Widths:").grid(row=3, column=0, sticky="w", pady=2); tk.Button(tab, text="Edit...", command=lambda: self.open_dict_editor("Column Widths", "column_id_widths")).grid(row=3, column=1, sticky="w")
        tk.Label(tab, text="Row Heights:").grid(row=4, column=0, sticky="w", pady=2); tk.Button(tab, text="Edit...", command=lambda: self.open_dict_editor("Row Heights", "row_heights")).grid(row=4, column=1, sticky="w")
        
    def create_footer_tab(self, sheet_conf):
        tab = ttk.Frame(self.notebook, padding=10)
        self.notebook.add(tab, text="Footer")
        footer_conf = sheet_conf.get("footer_configurations", {})
        
        tk.Label(tab, text="Total Text:").grid(row=0, column=0, sticky="w", pady=2); self.footer_total_text = tk.Entry(tab); self.footer_total_text.insert(0, footer_conf.get("total_text", "")); self.footer_total_text.grid(row=0, column=1, sticky="ew")
        tk.Label(tab, text="Total Text Column ID:").grid(row=1, column=0, sticky="w", pady=2); self.footer_total_text_id = tk.Entry(tab); self.footer_total_text_id.insert(0, footer_conf.get("total_text_column_id", "")); self.footer_total_text_id.grid(row=1, column=1, sticky="ew")
        tk.Label(tab, text="Pallet Count Column ID:").grid(row=2, column=0, sticky="w", pady=2); self.footer_pallet_id = tk.Entry(tab); self.footer_pallet_id.insert(0, footer_conf.get("pallet_count_column_id", "")); self.footer_pallet_id.grid(row=2, column=1, sticky="ew")
        tk.Label(tab, text="Sum Column IDs:").grid(row=3, column=0, sticky="w", pady=2); self.footer_sum_ids = tk.Entry(tab, width=80); self.footer_sum_ids.insert(0, ", ".join(footer_conf.get("sum_column_ids", []))); self.footer_sum_ids.grid(row=3, column=1, sticky="ew")
        tk.Label(tab, text="Number Formats:").grid(row=4, column=0, sticky="w", pady=2); tk.Button(tab, text="Edit...", command=lambda: self.open_dict_editor("Footer Number Formats", "number_formats")).grid(row=4, column=1, sticky="w")
        
    def create_advanced_tab(self, sheet_conf):
        tab = ttk.Frame(self.notebook, padding=10)
        self.notebook.add(tab, text="Advanced")
        
        tk.Label(tab, text="Weight Summary Config:").grid(row=0, column=0, sticky="w", pady=2); self.adv_weight_enabled = tk.BooleanVar(value=sheet_conf.get("weight_summary_config", {}).get("enabled", False)); tk.Checkbutton(tab, text="Enabled", variable=self.adv_weight_enabled).grid(row=0, column=1, sticky="w")
        tk.Label(tab, text="Label Col ID:").grid(row=1, column=0, sticky="w", pady=2); self.adv_weight_label = tk.Entry(tab); self.adv_weight_label.insert(0, sheet_conf.get("weight_summary_config", {}).get("label_col_id", "")); self.adv_weight_label.grid(row=1, column=1, sticky="ew")
        tk.Label(tab, text="Value Col ID:").grid(row=2, column=0, sticky="w", pady=2); self.adv_weight_value = tk.Entry(tab); self.adv_weight_value.insert(0, sheet_conf.get("weight_summary_config", {}).get("value_col_id", "")); self.adv_weight_value.grid(row=2, column=1, sticky="ew")
        
        ttk.Separator(tab, orient='horizontal').grid(row=3, column=0, columnspan=2, sticky='ew', pady=10)

        tk.Label(tab, text="Static Content Before Footer:").grid(row=4, column=0, sticky="w", pady=2); tk.Button(tab, text="Edit...", command=lambda: self.open_dict_editor("Static Content Before Footer", "static_content_before_footer")).grid(row=4, column=1, sticky="w")
        tk.Label(tab, text="Data Cell Merging Rule:").grid(row=5, column=0, sticky="w", pady=2); tk.Button(tab, text="Edit...", command=lambda: self.open_dict_editor("Data Cell Merging Rule", "data_cell_merging_rule")).grid(row=5, column=1, sticky="w")
        tk.Label(tab, text="Merge Rules Before Footer:").grid(row=6, column=0, sticky="w", pady=2); tk.Button(tab, text="Edit...", command=lambda: self.open_dict_editor("Merge Rules Before Footer", "merge_rules_before_footer")).grid(row=6, column=1, sticky="w")

    def open_dict_editor(self, title, data_key):
        initial_data = self.temp_dict_data.get(data_key, {})
        def on_save_callback(new_data):
            self.temp_dict_data[data_key] = new_data
        DictionaryEditor(self.root, title, initial_data, on_save_callback)

    def update_config_from_ui(self):
        sheet_name = self.selected_sheet_name.get()
        if not sheet_name or sheet_name not in self.config_data.get("data_mapping", {}): 
            return
        
        sheet_conf = self.config_data["data_mapping"][sheet_name]
        
        if self.sheet_start_row:
            try: sheet_conf["start_row"] = int(self.sheet_start_row.get())
            except (ValueError, AttributeError): sheet_conf["start_row"] = 21
            self.config_data["sheet_data_map"][sheet_name] = self.sheet_data_source.get()
            sheet_conf["summary"] = self.sheet_summary_var.get()
            sheet_conf["add_blank_before_footer"] = self.sheet_bbf_var.get()
        
        header_data = []
        for row_widgets in self.header_widgets:
            if row_widgets and row_widgets['text'].winfo_exists():
                try:
                    header_data.append({
                        "row": int(row_widgets['row'].get()), 
                        "col": int(row_widgets['col'].get()), 
                        "text": row_widgets['text'].get(), 
                        "id": row_widgets['id'].get(), 
                        "rowspan": int(row_widgets['rowspan'].get()), 
                        "colspan": int(row_widgets['colspan'].get())
                    })
                except (ValueError, tk.TclError): 
                    print(f"Skipping invalid or incomplete header row data for sheet '{sheet_name}'.")
        sheet_conf['header_to_write'] = header_data

        mappings_data = {}
        for widgets in self.mapping_widgets:
            if not widgets["name"].winfo_exists(): continue
            name = widgets["name"].get().strip()
            if not name: continue
            
            rule = {"id": widgets["id"].get()}
            if widgets["fallback"].get().strip(): rule["fallback_on_none"] = widgets["fallback"].get().strip()

            rule_type = widgets["type_combo"].get()
            if rule_type == "From Data Key":
                key_index_widget = widgets.get("key_index")
                if key_index_widget:
                    try: rule["key_index"] = int(key_index_widget.get())
                    except (ValueError, tk.TclError): pass
            elif rule_type == "From Data Value":
                value_key_widget = widgets.get("value_key")
                if value_key_widget: rule["value_key"] = value_key_widget.get()
            elif rule_type == "Formula":
                rule["type"] = "formula"
                rule["formula_template"] = widgets.get("formula_template", tk.Entry()).get()
                inputs_str = widgets.get("inputs", tk.Entry()).get()
                rule["inputs"] = [i.strip() for i in inputs_str.split(',') if i.strip()]
            
            mappings_data[name] = rule
        sheet_conf['mappings'] = mappings_data

        styling_conf = sheet_conf.get("styling", {})
        if hasattr(self, 'style_force_text'):
            styling_conf["force_text_format_ids"] = [i.strip() for i in self.style_force_text.get().split(',') if i.strip()]
            styling_conf["column_ids_with_full_grid"] = [i.strip() for i in self.style_full_grid.get().split(',') if i.strip()]
        for key in ['column_id_styles', 'column_id_widths', 'row_heights']:
            if key in self.temp_dict_data: styling_conf[key] = self.temp_dict_data[key]
        sheet_conf["styling"] = styling_conf

        footer_conf = sheet_conf.get("footer_configurations", {})
        if hasattr(self, 'footer_total_text'):
            footer_conf["total_text"] = self.footer_total_text.get()
            footer_conf["total_text_column_id"] = self.footer_total_text_id.get()
            footer_conf["pallet_count_column_id"] = self.footer_pallet_id.get()
            footer_conf["sum_column_ids"] = [i.strip() for i in self.footer_sum_ids.get().split(',') if i.strip()]
        if 'number_formats' in self.temp_dict_data: footer_conf["number_formats"] = self.temp_dict_data['number_formats']
        sheet_conf["footer_configurations"] = footer_conf
        
        weight_conf = sheet_conf.get("weight_summary_config", {})
        if hasattr(self, 'adv_weight_enabled'):
            weight_conf["enabled"] = self.adv_weight_enabled.get()
            weight_conf["label_col_id"] = self.adv_weight_label.get()
            weight_conf["value_col_id"] = self.adv_weight_value.get()
        sheet_conf["weight_summary_config"] = weight_conf
        
        for key in ['static_content_before_footer', 'data_cell_merging_rule', 'merge_rules_before_footer']:
            if key in self.temp_dict_data: sheet_conf[key] = self.temp_dict_data[key]

        self.config_data["data_mapping"][sheet_name] = sheet_conf

if __name__ == "__main__":
    root = tk.Tk()
    app = JsonConfigEditorApp(root)
    root.mainloop()