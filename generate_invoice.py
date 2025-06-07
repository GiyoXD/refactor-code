# generate_invoice.py
# Main script to orchestrate invoice generation using templates, config files, and utility functions.
# Updated: Handles multi-table processing and configurable header writing/finding.
# Updated: Added support for loading data from .pkl (Pickle) files.
# Updated: Removed fallback to find_header; requires write_header config for single-table sheets.
# Updated: Adds a merged blank row between multi-table sections (e.g., Packing List).
# Updated: Leaves a single blank row after the last table based on config 'row_spacing' (removed final merge).
# REVISED AGAIN: Corrected flag retrieval and passing in generate_invoice.py. Includes final debug prints.
# Updated: Added configurable merging for rows adjacent to header/footer.
# Updated: Integrated configurable styling (font, alignment, widths) from config file.
# Updated: Integrated pallet order tracking across multi-table chunks.
# MODIFIED: Calculates final_grand_total_pallets globally before sheet loop and passes it to all fill_invoice_data calls.

import os
import json
import pickle # Import pickle module
import argparse
import shutil
import openpyxl
import traceback
import sys
from pathlib import Path
from typing import Optional, Dict, Any, Union
import ast # <-- Add import for literal_eval
from decimal import Decimal # <-- Add import for Decimal evaluation
import re # <-- Add import for regular expressions
from openpyxl.utils import get_column_letter # REMOVED range_boundaries
import text_replace_utils # Ensure this is imported

# --- Import utility functions ---
try:
    # Ensure invoice_utils.py corresponds to the latest version with pallet order updates
    import invoice_utils
    import merge_utils # <-- Import the new merge utility module
    print("Successfully imported invoice_utils and merge_utils.")
except ImportError as import_err:
    print("------------------------------------------------------")
    print(f"FATAL ERROR: Could not import required utility modules: {import_err}")
    print("Please ensure invoice_utils.py and merge_utils.py are in the same directory as generate_invoice.py.")
    print("------------------------------------------------------")
    sys.exit(1)

# --- Helper Functions (derive_paths, load_config, load_data) ---
# Assume these functions exist as previously defined in the uploaded file.
# They are omitted here for brevity but are required for the script to work.
# Make sure to include the full code for these functions from your original file.
# --- Placeholder for Required Helper Functions ---
# NOTE: Replace 'pass' with the actual function definitions from your original file
def derive_paths(input_data_path_str: str, template_dir_str: str, config_dir_str: str) -> Optional[Dict[str, Path]]:
    """
    Derives template and config file paths based on the input data filename.
    Checks if template/config paths are valid directories.
    Assumes data file is named like TEMPLATE_NAME.xxx or TEMPLATE_NAME_data.xxx
    Attempts prefix matching if exact match is not found.
    """
    print(f"Deriving paths from input: {input_data_path_str}")
    try:
        input_data_path = Path(input_data_path_str).resolve()
        template_dir = Path(template_dir_str).resolve()
        config_dir = Path(config_dir_str).resolve()

        if not input_data_path.is_file(): print(f"Error: Input data file not found: {input_data_path}"); return None
        if not template_dir.is_dir(): print(f"Error: Template directory not found: {template_dir}"); return None
        if not config_dir.is_dir(): print(f"Error: Config directory not found: {config_dir}"); return None

        base_name = input_data_path.stem
        template_name_part = base_name
        suffixes_to_remove = ['_data', '_input', '_pkl']
        prefixes_to_remove = ['data_']

        for suffix in suffixes_to_remove:
            if base_name.lower().endswith(suffix):
                template_name_part = base_name[:-len(suffix)]
                break
        else: # Only check prefixes if no suffix was removed
            for prefix in prefixes_to_remove:
                if base_name.lower().startswith(prefix):
                    template_name_part = base_name[len(prefix):]
                    break

        if not template_name_part:
            print(f"Error: Could not derive template name part from: '{base_name}'")
            return None
        print(f"Derived initial template name part: '{template_name_part}'")

        # --- Attempt 1: Exact Match ---
        exact_template_filename = f"{template_name_part}.xlsx"
        exact_config_filename = f"{template_name_part}_config.json"
        exact_template_path = template_dir / exact_template_filename
        exact_config_path = config_dir / exact_config_filename
        print(f"Checking for exact match: Template='{exact_template_path}', Config='{exact_config_path}'")

        if exact_template_path.is_file() and exact_config_path.is_file():
            print("Found exact match for template and config.")
            return {"data": input_data_path, "template": exact_template_path, "config": exact_config_path}
        else:
            print("Exact match not found. Attempting prefix matching...")

            # --- Attempt 2: Prefix Match ---
            prefix_match = re.match(r'^([a-zA-Z]+)', template_name_part) # Extract leading letters
            if prefix_match:
                prefix = prefix_match.group(1)
                print(f"Extracted prefix: '{prefix}'")
                prefix_template_filename = f"{prefix}.xlsx"
                prefix_config_filename = f"{prefix}_config.json"
                prefix_template_path = template_dir / prefix_template_filename
                prefix_config_path = config_dir / prefix_config_filename
                print(f"Checking for prefix match: Template='{prefix_template_path}', Config='{prefix_config_path}'")

                if prefix_template_path.is_file() and prefix_config_path.is_file():
                    print("Found prefix match for template and config.")
                    return {"data": input_data_path, "template": prefix_template_path, "config": prefix_config_path}
                else:
                    print("Prefix match not found.")
            else:
                print("Could not extract a letter-based prefix.")

            # --- No Match Found ---
            print(f"Error: Could not find matching template/config files using exact ('{template_name_part}') or prefix methods.")
            # Report specific missing files based on the exact match attempt
            if not exact_template_path.is_file(): print(f"Error: Template file not found: {exact_template_path}")
            if not exact_config_path.is_file(): print(f"Error: Configuration file not found: {exact_config_path}")
            return None

    except Exception as e:
        print(f"Error deriving file paths: {e}")
        traceback.print_exc()
        return None

def load_config(config_path: Path) -> Optional[Dict[str, Any]]:
    """Loads and parses the JSON configuration file."""
    print(f"Loading configuration from: {config_path}")
    try:
        with open(config_path, 'r', encoding='utf-8') as f: config_data = json.load(f)
        print("Configuration loaded successfully.")
        if not isinstance(config_data, dict): print("Error: Config file is not a valid JSON object."); return None
        # Basic validation (add checks for 'styling' if required globally, but usually per-sheet)
        required_keys = ['sheets_to_process', 'sheet_data_map', 'footer_rules', 'data_mapping']
        missing_keys = [key for key in required_keys if key not in config_data]
        if missing_keys: print(f"Error: Config file missing required keys: {', '.join(missing_keys)}"); return None
        if not isinstance(config_data.get('data_mapping'), dict): print(f"Error: 'data_mapping' section is not a valid dictionary."); return None
        return config_data
    except json.JSONDecodeError as e: print(f"Error: Invalid JSON in configuration file {config_path}: {e}"); return None
    except Exception as e: print(f"Error loading configuration file {config_path}: {e}"); traceback.print_exc(); return None

def load_data(data_path: Path) -> Optional[Dict[str, Any]]:
    """ Loads and parses the input data file. Supports .json and .pkl. """
    print(f"Loading data from: {data_path}")
    invoice_data = None; file_suffix = data_path.suffix.lower()
    try:
        if file_suffix == '.json':
            print("Detected .json file...")
            with open(data_path, 'r', encoding='utf-8') as f: invoice_data = json.load(f)
            print("JSON data loaded successfully.")
        elif file_suffix == '.pkl':
            print("Detected .pkl file...");
            with open(data_path, 'rb') as f: invoice_data = pickle.load(f)
            print("Pickle data loaded successfully.")
        else: print(f"Error: Unsupported data file extension: '{file_suffix}'."); return None
        if not isinstance(invoice_data, dict): print("Error: Loaded data is not a dictionary."); return None

        # --- START AGGREGATION KEY CONVERSION ---
        # Use "initial_standard_aggregation" as requested
        aggregation_data_raw = invoice_data.get("standard_aggregation_results")
        if isinstance(aggregation_data_raw, dict):
            print("DEBUG: Found 'standard_aggregation_results'. Converting string keys to tuples...")
            aggregation_data_processed = {}
            converted_count = 0
            conversion_errors = 0
            # Regex to find Decimal('...') and capture the inner number string
            decimal_pattern = re.compile(r"Decimal\('(-?\d*\.?\d+)'\)") # Handles optional -, digits, optional decimal point

            for key_str, value_dict in aggregation_data_raw.items():
                processed_key_str = key_str # Initialize for error message
                try:
                    # Preprocess the string: Replace Decimal('...') with just the number string '...'
                    processed_key_str = decimal_pattern.sub(r"'\1'", key_str) # Replace with the number in quotes

                    # Now evaluate the processed string which should only contain literals
                    key_tuple = ast.literal_eval(processed_key_str)

                    # --- START MODIFIED POST-PROCESSING ---
                    # Convert tuple elements: Keep PO (idx 0) and Item (idx 1) as strings,
                    # convert Unit Price (idx 2) to float.
                    final_key_list = []
                    if isinstance(key_tuple, tuple) and len(key_tuple) >= 3:
                        # PO Number (Index 0): Keep as string
                        final_key_list.append(str(key_tuple[0]))

                        # Item Number (Index 1): Keep as string
                        final_key_list.append(str(key_tuple[1]))

                        # Unit Price (Index 2): Convert to float
                        unit_price_val = key_tuple[2]
                        if isinstance(unit_price_val, (int, float)):
                            final_key_list.append(float(unit_price_val))
                        elif isinstance(unit_price_val, str):
                            try:
                                final_key_list.append(float(unit_price_val))
                            except ValueError:
                                print(f"Warning: Could not convert unit price string '{unit_price_val}' to float for key '{key_str}'. Keeping as string.")
                                final_key_list.append(unit_price_val) # Keep original string on error
                        else:
                            # Handle other types if necessary, maybe try converting to float
                            try: final_key_list.append(float(unit_price_val))
                            except (ValueError, TypeError):
                                print(f"Warning: Could not convert unit price type '{type(unit_price_val)}' ({unit_price_val}) to float for key '{key_str}'. Keeping original type.")
                                final_key_list.append(unit_price_val) # Keep original on error

                        # Add any remaining elements from the original tuple (if any)
                        if len(key_tuple) > 3:
                            final_key_list.extend(key_tuple[3:])

                    else:
                        # Handle cases where the tuple doesn't have the expected structure
                        print(f"Warning: Evaluated key tuple '{key_tuple}' does not have expected length >= 3. Using original items.")
                        final_key_list = list(key_tuple) # Use original items

                    final_key_tuple = tuple(final_key_list)
                    # --- END MODIFIED POST-PROCESSING ---


                    if isinstance(final_key_tuple, tuple):
                        aggregation_data_processed[final_key_tuple] = value_dict
                        converted_count += 1
                    else:
                        # This case should be less likely now with the explicit tuple check above
                        print(f"Warning: Final key is not a tuple for processed key string '{processed_key_str}'. Original: '{key_str}'. Result: {final_key_tuple}")
                        conversion_errors += 1
                except (ValueError, SyntaxError, NameError, TypeError) as e:
                    print(f"Warning: Could not convert aggregation key string '{key_str}' (processed: '{processed_key_str}') to tuple: {e}")
                    conversion_errors += 1
            # Replace the original string-keyed dict with the tuple-keyed one
            # Update the key used for replacement as well
            invoice_data["standard_aggregation_results"] = aggregation_data_processed
            print(f"DEBUG: Finished key conversion. Converted: {converted_count}, Errors: {conversion_errors}")
        # --- END AGGREGATION KEY CONVERSION ---

        # --- START CUSTOM AGGREGATION KEY CONVERSION ---
        # Added block to handle custom_aggregation_results
        custom_aggregation_data_raw = invoice_data.get("custom_aggregation_results")
        if isinstance(custom_aggregation_data_raw, dict):
            print("DEBUG: Found 'custom_aggregation_results'. Converting string keys to tuples...")
            custom_aggregation_data_processed = {}
            custom_converted_count = 0
            custom_conversion_errors = 0
            # Reuse the same regex pattern
            decimal_pattern = re.compile(r"Decimal\('(-?\d*\.?\d+)'\)")

            for key_str, value_dict in custom_aggregation_data_raw.items():
                processed_key_str = key_str
                try:
                    processed_key_str = decimal_pattern.sub(r"'\1'", key_str)
                    key_tuple = ast.literal_eval(processed_key_str)

                    # Apply the same post-processing as standard aggregation if needed
                    # (Assuming the structure PO, Item, [Optional Price] is consistent)
                    final_key_list = []
                    if isinstance(key_tuple, tuple) and len(key_tuple) >= 2: # Custom might only have PO, Item
                        # PO Number (Index 0): Keep as string
                        final_key_list.append(str(key_tuple[0]))
                        # Item Number (Index 1): Keep as string
                        final_key_list.append(str(key_tuple[1]))
                        # Keep remaining elements (e.g., None in the example)
                        if len(key_tuple) > 2:
                            final_key_list.extend(key_tuple[2:])
                    else:
                        print(f"Warning: Custom key tuple '{key_tuple}' doesn't have expected length >= 2. Using original items.")
                        final_key_list = list(key_tuple)

                    final_key_tuple = tuple(final_key_list)

                    if isinstance(final_key_tuple, tuple):
                        custom_aggregation_data_processed[final_key_tuple] = value_dict
                        custom_converted_count += 1
                    else:
                        print(f"Warning: Final custom key is not a tuple for processed key string '{processed_key_str}'. Original: '{key_str}'. Result: {final_key_tuple}")
                        custom_conversion_errors += 1
                except (ValueError, SyntaxError, NameError, TypeError) as e:
                    print(f"Warning: Could not convert custom aggregation key string '{key_str}' (processed: '{processed_key_str}') to tuple: {e}")
                    custom_conversion_errors += 1

            invoice_data["custom_aggregation_results"] = custom_aggregation_data_processed
            print(f"DEBUG: Finished key conversion for custom_aggregation_results. Converted: {custom_converted_count}, Errors: {custom_conversion_errors}")
        # --- END CUSTOM AGGREGATION KEY CONVERSION ---

        return invoice_data
    except json.JSONDecodeError as e: print(f"Error: Invalid JSON in data file {data_path}: {e}"); return None
    except pickle.UnpicklingError as e: print(f"Error: Could not unpickle data file {data_path}: {e}"); return None
    except FileNotFoundError: print(f"Error: Data file not found at {data_path}"); return None
    except Exception as e: print(f"Error loading data file {data_path}: {e}"); traceback.print_exc(); return None
# --- End Placeholder ---


# --- Main Orchestration Logic ---
def main():
    """Main function to orchestrate invoice generation."""
    parser = argparse.ArgumentParser(description="Generate Invoice from Template and Data using configuration files.")
    parser.add_argument("input_data_file", help="Path to the input data file (.json or .pkl). Filename base determines template/config.")
    parser.add_argument("-o", "--output", default="result.xlsx", help="Path for the output Excel file (default: result.xlsx)")
    parser.add_argument("-t", "--templatedir", default="./TEMPLATE", help="Directory containing template Excel files (default: ./TEMPLATE)")
    parser.add_argument("-c", "--configdir", default="./configs", help="Directory containing configuration JSON files (default: ./configs)")
    parser.add_argument("--fob", action="store_true", help="Generate FOB version using final_fob_compounded_result for Invoice/Contract sheets.")
    parser.add_argument("--custom", action="store_true", help="Enable custom processing logic (details TBD).")
    args = parser.parse_args()

    print("--- Starting Invoice Generation ---")
    print(f"Input Data: {args.input_data_file}"); print(f"Template Dir: {args.templatedir}"); print(f"Config Dir: {args.configdir}"); print(f"Output File: {args.output}")

    print("\n1. Deriving file paths..."); paths = derive_paths(args.input_data_file, args.templatedir, args.configdir)
    if not paths: sys.exit(1)

    print("\n2. Loading configuration and data..."); config = load_config(paths['config']); invoice_data = load_data(paths['data'])
    if not config or not invoice_data: sys.exit(1)

    print(f"\n3. Copying template '{paths['template'].name}' to '{args.output}'..."); output_path = Path(args.output).resolve()
    try:
        output_path.parent.mkdir(parents=True, exist_ok=True);
        shutil.copy(paths['template'], output_path);
    except Exception as e:
        print(f"Error copying template: {e}"); sys.exit(1)
    print(f"Template copied successfully to {output_path}")

    print("\n4. Processing workbook...");
    workbook = None; processing_successful = True

    try:
        workbook = openpyxl.load_workbook(output_path)

        # --- Determine sheets to process ---
        sheets_to_process_config = config.get('sheets_to_process', [])
        if not sheets_to_process_config:
            sheets_to_process = [workbook.active.title] if workbook.active else []
        else:
            sheets_to_process = [s for s in sheets_to_process_config if s in workbook.sheetnames] # Filter valid sheets

        if not sheets_to_process:
            print("Error: No valid sheets found or specified to process.")
            if workbook:
                try: workbook.close()
                except Exception: pass
            sys.exit(1) # Exit if no sheets to process

        # --- Store Original Merges BEFORE processing using merge_utils ---
        original_merges = merge_utils.store_original_merges(workbook, sheets_to_process)
        print("DEBUG: Stored original merges structure:")

        # --- Get other config sections ---
        sheet_data_map = config.get('sheet_data_map', {})
        global_footer_rules = config.get('footer_rules', {})
        data_mapping_config = config.get('data_mapping', {})

        # ***** MOVED GLOBAL PALLET CALCULATION HERE *****
        final_grand_total_pallets = 0
        print("DEBUG: Pre-calculating final grand total pallets globally...")
        processed_tables_data_for_calc = invoice_data.get('processed_tables_data', {})
        if isinstance(processed_tables_data_for_calc, dict) and processed_tables_data_for_calc:
            temp_total = 0
            table_keys_for_calc = sorted(processed_tables_data_for_calc.keys(), key=lambda x: int(x) if str(x).isdigit() else float('inf'))
            for temp_key in table_keys_for_calc:
                temp_table_data = processed_tables_data_for_calc.get(str(temp_key))
                if isinstance(temp_table_data, dict):
                    pallet_counts = temp_table_data.get("pallet_count", [])
                    if isinstance(pallet_counts, list):
                        for count in pallet_counts:
                            try:
                                temp_total += int(count)
                            except (ValueError, TypeError):
                                pass # Ignore non-integer counts
            final_grand_total_pallets = temp_total
        else:
            print("DEBUG: 'processed_tables_data' not found or empty in input data. final_grand_total_pallets remains 0.")
        print(f"DEBUG: Globally calculated final grand total pallets: {final_grand_total_pallets}")
        # ***** END GLOBAL PALLET CALCULATION *****


        print(f"\nWill process sheets: {sheets_to_process}")
        # --- Start Sheet Processing Loop ---
        for sheet_name in sheets_to_process:
            print(f"\n--- Processing Sheet: '{sheet_name}' ---")
            if sheet_name not in workbook.sheetnames:
                print(f"Warning: Sheet '{sheet_name}' not found at processing time. Skipping.")
                continue
            worksheet = workbook[sheet_name]

            # --- Get sheet-specific config sections ---
            sheet_mapping_section = data_mapping_config.get(sheet_name, {}) # Use .get for safety
            data_source_indicator = sheet_data_map.get(sheet_name) # Get indicator from config

            # --- Check for FOB flag override ---
            if args.fob and sheet_name in ["Invoice", "Contract"]:
                print(f"DEBUG: --fob flag active. Overriding data source for '{sheet_name}' to 'fob_aggregation'.")
                data_source_indicator = 'fob_aggregation'
            # --- End FOB flag override ---

            sheet_styling_config = sheet_mapping_section.get("styling") # Get styling rules dict or None

            if not sheet_mapping_section: print(f"Warning: No 'data_mapping' section for sheet '{sheet_name}'. Skipping."); continue
            if not data_source_indicator: print(f"Warning: No 'sheet_data_map' entry for sheet '{sheet_name}' (or FOB override failed). Skipping."); continue # Adjusted warning

            # --- Retrieve flags and mappings ONCE per sheet ---
            add_blank_after_hdr_flag = sheet_mapping_section.get("add_blank_after_header", False)
            static_content_after_hdr_dict = sheet_mapping_section.get("static_content_after_header", {})
            add_blank_before_ftr_flag = sheet_mapping_section.get("add_blank_before_footer", False)
            static_content_before_ftr_dict = sheet_mapping_section.get("static_content_before_footer", {})
            sheet_inner_mapping_rules_dict = sheet_mapping_section.get('mappings', {})
            final_row_spacing = sheet_mapping_section.get('row_spacing', 0)
            merge_rules_after_hdr = sheet_mapping_section.get("merge_rules_after_header", {})
            merge_rules_before_ftr = sheet_mapping_section.get("merge_rules_before_footer", {})
            merge_rules_footer = sheet_mapping_section.get("merge_rules_footer", {})
            data_cell_merging_rules = sheet_mapping_section.get("data_cell_merging_rule", None)

            print(f"DEBUG Check Flags Read for Sheet '{sheet_name}': after_hdr={add_blank_after_hdr_flag}, before_ftr={add_blank_before_ftr_flag}")
            if sheet_styling_config: print("DEBUG: Styling config found for this sheet.")
            else: print("DEBUG: No styling config found for this sheet.")

            all_tables_data = invoice_data.get('processed_tables_data', {})
            table_keys = sorted(all_tables_data.keys(), key=lambda x: int(x) if str(x).isdigit() else float('inf'))
            # ================================================================
            # --- Handle Multi-Table Case (e.g., Packing List) ---
            # ================================================================
            if data_source_indicator == "processed_tables_multi":
                print(f"Processing sheet '{sheet_name}' as multi-table (write header mode).")
                all_tables_data = invoice_data.get('processed_tables_data', {})
                if not all_tables_data or not isinstance(all_tables_data, dict): print(f"Warning: 'processed_tables_data' not found/valid. Skipping '{sheet_name}'."); continue

                header_to_write = sheet_mapping_section.get('header_to_write');
                header_merge_rules = sheet_mapping_section.get('header_merge_rules')
                start_row = sheet_mapping_section.get('start_row') # Use config start_row
                if not start_row or not header_to_write: print(f"Error: Config for multi-table '{sheet_name}' missing 'start_row' or 'header_to_write'. Skipping."); processing_successful = False; continue

                table_keys = sorted(all_tables_data.keys(), key=lambda x: int(x) if str(x).isdigit() else float('inf'))
                print(f"Found table keys in data: {table_keys}"); num_tables = len(table_keys); last_table_header_info = None

                # --- V11: Pre-calculate total rows needed for bulk insert --- 
                total_rows_to_insert = 0
                print("--- Pre-calculating total rows for multi-table section ---")
                for i, table_key in enumerate(table_keys):
                    table_data_to_fill = all_tables_data.get(str(table_key))
                    if not table_data_to_fill or not isinstance(table_data_to_fill, dict): continue # Skip if no data

                    # Rows for header
                    num_header_rows = len(header_to_write) if isinstance(header_to_write, list) else 0
                    total_rows_to_insert += num_header_rows
                    print(f"  Table {table_key}: +{num_header_rows} (header)")

                    # Rows for blank after header (if applicable)
                    if add_blank_after_hdr_flag: total_rows_to_insert += 1; print(f"  Table {table_key}: +1 (blank after header)")

                    # Rows for data
                    num_data_rows = 0
                    if isinstance(table_data_to_fill, dict):
                        max_len = 0
                        for value in table_data_to_fill.values():
                            if isinstance(value, list): max_len = max(max_len, len(value))
                        num_data_rows = max_len
                    total_rows_to_insert += num_data_rows
                    print(f"  Table {table_key}: +{num_data_rows} (data rows)")

                    # Rows for blank before footer (if applicable)
                    if add_blank_before_ftr_flag: total_rows_to_insert += 1; print(f"  Table {table_key}: +1 (blank before footer)")

                    # Row for footer
                    total_rows_to_insert += 1 # Always add 1 for the footer row itself
                    print(f"  Table {table_key}: +1 (footer)")

                    # Row for spacer (if not the last table)
                    if i < num_tables - 1: total_rows_to_insert += 1; print(f"  Table {table_key}: +1 (spacer)")

                # Rows for Grand Total (if more than one table)
                if num_tables > 1: total_rows_to_insert += 1; print("  Overall: +1 (Grand Total Row)")

                # Rows for Summary flag
                summary_flag = sheet_mapping_section.get("summary", False)
                if summary_flag and num_tables > 0: # Only add if GT row was added
                    total_rows_to_insert += 2; print("  Overall: +2 (Summary Flag Rows)")

                # Rows for Final Spacing
                if final_row_spacing > 0: total_rows_to_insert += final_row_spacing; print(f"  Overall: +{final_row_spacing} (Final Spacing)")

                print(f"--- Total rows to insert for multi-table section: {total_rows_to_insert} ---")
                # --- End Pre-calculation ---

                # --- V11: Perform Bulk Insert --- 
                if total_rows_to_insert > 0:
                    try:
                        print(f"Inserting {total_rows_to_insert} rows at index {start_row} for sheet '{sheet_name}'...")
                        worksheet.insert_rows(start_row, amount=total_rows_to_insert)
                        # Unmerge the inserted block to prevent issues
                        invoice_utils.safe_unmerge_block(worksheet, start_row, start_row + total_rows_to_insert - 1, worksheet.max_column)
                        print("Bulk rows inserted and unmerged successfully.")
                    except Exception as bulk_insert_err:
                        print(f"ERROR: Failed bulk row insert for multi-table: {bulk_insert_err}")
                        processing_successful = False
                        # Decide how to handle this - skip sheet?
                        continue # Skip processing this sheet if bulk insert fails
                # --- End Bulk Insert ---

                # --- V11: Initialize write pointer --- 
                write_pointer_row = start_row # Start writing at the beginning of the inserted block
 
                # ***** INITIALIZE GRAND TOTAL FOR *THIS SHEET TYPE* & PALLET ORDER VARIABLES *****
                grand_total_pallets_for_summary_row = 0
                all_data_ranges = [] # List to store tuples of (start_row, end_row) for SUM
                # ***** END INITIALIZE *****
 
                # ***** REMOVED REDUNDANT PRE-CALCULATION LOOP - NOW DONE GLOBALLY *****
 
                # --- V11: Main loop now only writes data, doesn't insert --- # TODO
                for i, table_key in enumerate(table_keys):
                    print(f"\nProcessing table key: '{table_key}' ({i+1}/{num_tables})")
                    table_data_to_fill = all_tables_data.get(str(table_key))
                    if not table_data_to_fill or not isinstance(table_data_to_fill, dict): print(f"Warning: No/invalid data for table key '{table_key}'. Skipping."); continue
 
                    print(f"Writing header for table '{table_key}' at row {write_pointer_row}...");
                    written_header_info = invoice_utils.write_header(
                        worksheet, write_pointer_row, header_to_write, header_merge_rules, sheet_styling_config
                    )
                    if not written_header_info: print(f"Error writing header for table '{table_key}'. Skipping sheet."); processing_successful = False; break
                    last_table_header_info = written_header_info # Keep track for width setting later
 
                    # Update write pointer after header
                    num_header_rows = len(header_to_write) if isinstance(header_to_write, list) else 0
                    write_pointer_row += num_header_rows
 
                    print(f"Filling data and footer for table '{table_key}' starting near row {write_pointer_row}...")
                    # Pass the current write pointer as the effective 'start row' for fill_invoice_data
                    # It will write header, data, footer starting from here
                    # NOTE: We need to adjust fill_invoice_data to use the passed start row correctly
                    #       instead of header_info['second_row_index'] + 1
 
                    # Modify the header_info dict passed to fill_invoice_data dynamically
                    temp_header_info = written_header_info.copy()
                    temp_header_info['first_row_index'] = write_pointer_row - num_header_rows # The row where header started
                    temp_header_info['second_row_index'] = write_pointer_row - 1 # The last row of the header
 
                    fill_success, next_row_after_chunk, data_start, data_end, table_pallets = invoice_utils.fill_invoice_data(
                        worksheet=worksheet,
                        sheet_name=sheet_name,
                        sheet_config=sheet_mapping_section, # Pass current sheet's config
                        all_sheet_configs=data_mapping_config, # <--- Pass the full config map
                        data_source=table_data_to_fill,
                        data_source_type='processed_tables',
                        header_info=temp_header_info,
                        mapping_rules=sheet_inner_mapping_rules_dict,
                        sheet_styling_config=sheet_styling_config,
                        add_blank_after_header=add_blank_after_hdr_flag,
                        static_content_after_header=static_content_after_hdr_dict,
                        add_blank_before_footer=add_blank_before_ftr_flag,
                        static_content_before_footer=static_content_before_ftr_dict,
                        merge_rules_after_header=merge_rules_after_hdr,
                        merge_rules_before_footer=merge_rules_before_ftr,
                        merge_rules_footer=merge_rules_footer,
                        footer_info=None, max_rows_to_fill=None,
                        grand_total_pallets=final_grand_total_pallets,
                        custom_flag=args.custom,
                        data_cell_merging_rules=data_cell_merging_rules,
                    )
                    # fill_invoice_data now handles writing blank rows, data, footer row
                    # within the allocated space. next_row_after_chunk is the row AFTER its footer.
 
                    if fill_success:
                        print(f"Finished table '{table_key}'. Next available write pointer is {next_row_after_chunk}")
                        grand_total_pallets_for_summary_row += table_pallets
                        if data_start > 0 and data_end >= data_start: all_data_ranges.append((data_start, data_end))
 
                        write_pointer_row = next_row_after_chunk # Update pointer to be after the chunk
 
                        is_last_table = (i == num_tables - 1)
                        if not is_last_table: 
                            # Write the spacer row content (optional, could just be blank)
                            spacer_row = write_pointer_row
                            num_cols_spacer = written_header_info.get('num_columns', worksheet.max_column)
                            if num_cols_spacer > 0:
                                try:
                                    print(f"Writing merged spacer row at {spacer_row} across {num_cols_spacer} columns...")
                                    # No insert needed, just merge and maybe clear/style
                                    invoice_utils.unmerge_row(worksheet, spacer_row, num_cols_spacer) # Ensure clear
                                    worksheet.merge_cells(start_row=spacer_row, start_column=1, end_row=spacer_row, end_column=num_cols_spacer)
                                    # Optionally add styling or blank value to the merged cell
                                    # worksheet.cell(row=spacer_row, column=1).border = ... 
                                    write_pointer_row += 1 # Advance pointer past the spacer row
                                except Exception as merge_err: 
                                    print(f"Warning: Failed to write/merge spacer row {spacer_row}: {merge_err}"); 
                                    write_pointer_row += 1 # Still advance pointer even if merge fails
                            else: 
                                print("Warning: Cannot determine table width for spacer.");
                                write_pointer_row += 1 # Advance pointer anyway
                        # No 'else' needed, pointer is already correct if it's the last table
                    else: 
                        print(f"Error filling data/footer for table '{table_key}'. Stopping."); 
                        processing_successful = False; break
                # --- End Table Loop ---
 
                # --- DEBUG: Check conditions before adding Grand Total Row --- 
                print("\n--- Checking conditions for Grand Total Row ---")
                print(f"  - processing_successful: {processing_successful}")
                print(f"  - last_table_header_info is not None: {last_table_header_info is not None}")
                if last_table_header_info:
                     print(f"  - last_table_header_info content (keys): {list(last_table_header_info.keys())}")
                print(f"  - num_tables > 1: {num_tables > 1} (num_tables={num_tables})") # Only add if more than 1 table
                # --- END DEBUG ---
 
                # ***** ADD GRAND TOTAL ROW (for multi-table summary) *****
                # V11: No insert needed, just write at the current pointer
                if processing_successful and last_table_header_info and num_tables > 1:
                    grand_total_row_num = write_pointer_row # Write at the current position
                    print(f"\n--- Adding Grand Total Row at index {grand_total_row_num} ---")
                    try:
                        # worksheet.insert_rows(grand_total_row_num) # REMOVED INSERT
                        invoice_utils.unmerge_row(worksheet, grand_total_row_num, last_table_header_info['num_columns'])
                         
                        # --- Define Grand Total Styling (using header config) ---
                        grand_total_font = invoice_utils.bold_font # Default fallback
                        grand_total_row_height = None # Default fallback (no specific height)
                        if sheet_styling_config:
                            header_font_config = sheet_styling_config.get("header_font")
                            if header_font_config and isinstance(header_font_config, dict):
                                try:
                                    grand_total_font = openpyxl.styles.Font(
                                        name=header_font_config.get("name", "Calibri"),
                                        size=header_font_config.get("size", 11),
                                        bold=header_font_config.get("bold", True),
                                        italic=header_font_config.get("italic", False),
                                    )
                                except Exception as font_err:
                                    print(f"Warning: Error creating font from header_font config: {font_err}. Using default bold font.")
                            row_heights_config = sheet_styling_config.get("row_heights")
                            if row_heights_config and isinstance(row_heights_config, dict):
                                grand_total_row_height = row_heights_config.get("header")

                        if grand_total_row_height is not None:
                            try: worksheet.row_dimensions[grand_total_row_num].height = float(grand_total_row_height)
                            except (ValueError, TypeError): print(f"Warning: Invalid header row height value '{grand_total_row_height}' in config.")

                        # Get necessary info from last header/styling
                        column_map = last_table_header_info['column_map']
                        idx_to_header_map = {v: k for k, v in column_map.items()}
                        num_cols = last_table_header_info['num_columns']

                        # --- Write "TOTAL OF:" text --- (Changed from GRAND TOTAL)
                        po_col_idx = column_map.get("P.O Nº") or column_map.get("P.O N°") or column_map.get("P.O N °") 
                        total_text_col_idx = po_col_idx # Default to PO column
                        # Fallback logic if PO not found (find first text-like column)
                        if not total_text_col_idx:
                            for c_idx in range(1, num_cols + 1):
                                hdr = idx_to_header_map.get(c_idx, "").lower()
                                if "pallet" in hdr or "po" in hdr or "item" in hdr: # Add more keywords if needed
                                    total_text_col_idx = c_idx
                                    break
                            if not total_text_col_idx: total_text_col_idx = 2 # Absolute fallback

                        if total_text_col_idx:
                            gt_cell = worksheet.cell(row=grand_total_row_num, column=total_text_col_idx, value="TOTAL OF: ")
                            gt_cell.font = grand_total_font
                            invoice_utils._apply_cell_style(gt_cell, idx_to_header_map.get(total_text_col_idx), sheet_styling_config)
                            print(f"DEBUG: Wrote 'TOTAL OF:' to column {total_text_col_idx}")
                        else: print("Warning: Could not find suitable column for 'TOTAL OF:'.")


                        # --- Write Grand Total Pallets (for summary row) ---
# --- Write Grand Total Pallets (for summary row - Configurable Column) ---
                        gt_pallet_header_config_name = sheet_mapping_section.get("footer_pallet_count_column_header")
                        gt_pallet_col_idx = None

                        # Attempt 1: Use configured header name
                        if gt_pallet_header_config_name:
                            gt_pallet_col_idx = column_map.get(gt_pallet_header_config_name)
                            if gt_pallet_col_idx is None:
                                print(f"Warning: Configured Grand Total Row pallet header '{gt_pallet_header_config_name}' not found in column_map. Falling back...")
                        else:
                            print("DEBUG: 'grand_total_row_pallet_column_header' not configured for this sheet. Falling back for pallet count column.")

                        # Attempt 2: Fallback to "Description" header
                        if gt_pallet_col_idx is None:
                            gt_pallet_col_idx = column_map.get("Description")
                            if gt_pallet_col_idx is None:
                                print("Warning: 'Description' header not found for Grand Total Row pallet count. Trying further fallback...")                        
                        desc_col_idx = gt_pallet_col_idx
                        if desc_col_idx:
                            pallet_cell = worksheet.cell(row=grand_total_row_num, column=desc_col_idx)
                            # Use the locally accumulated grand_total_pallets_for_summary_row here
                            pallet_cell.value = f"{grand_total_pallets_for_summary_row} PALLETS"
                            pallet_cell.font = grand_total_font
                            invoice_utils._apply_cell_style(pallet_cell, idx_to_header_map.get(desc_col_idx), sheet_styling_config)
                        else: print("Warning: Could not find Description column for grand total pallets summary row.")

                        # --- Write Grand Total SUM Formulas ---
                        headers_to_sum = ["PCS", "SF", "N.W (kgs)", "G.W (kgs)", "CBM"] # Adjust as needed
                        if all_data_ranges: # Only sum if there were valid data ranges
                            print(f"DEBUG: Creating grand total SUM formulas using ranges: {all_data_ranges}")
                            for header_name in headers_to_sum:
                                col_idx = column_map.get(header_name)
                                if col_idx:
                                    col_letter = get_column_letter(col_idx)
                                    sum_parts = [f"{col_letter}{start}:{col_letter}{end}" for start, end in all_data_ranges]
                                    formula = f"=SUM({','.join(sum_parts)})"
                                    try:
                                        sum_cell = worksheet.cell(row=grand_total_row_num, column=col_idx, value=formula)
                                        sum_cell.font = grand_total_font
                                        invoice_utils._apply_cell_style(sum_cell, header_name, sheet_styling_config)
                                        print(f"DEBUG: Wrote SUM formula '{formula}' for '{header_name}'")
                                    except Exception as e: print(f"Error writing SUM formula for {header_name}: {e}")
                                else: print(f"Warning: Header '{header_name}' not found for grand total SUM.")
                        else: print("DEBUG: No valid data ranges found. Skipping grand total SUM formulas.")

                        # --- Apply Styling/Borders to Grand Total Row ---
                        print(f"DEBUG: Styling grand total row {grand_total_row_num}...")
                        for c_idx in range(1, num_cols + 1):
                            cell = worksheet.cell(row=grand_total_row_num, column=c_idx)
                            # cell.border = invoice_utils.thin_border # Apply full border
                            # REMEMBER to uncomment this line when you're ready to apply the full border
                            invoice_utils._apply_cell_style(cell, idx_to_header_map.get(c_idx), sheet_styling_config)
                            cell.font = grand_total_font # Ensure font override
 
                        write_pointer_row += 1 # Update write pointer position after GT row
                        print(f"--- Finished Adding Grand Total Row. Next write pointer: {write_pointer_row} ---")
 
                    except Exception as gt_err:
                        print(f"--- ERROR adding Grand Total row: {gt_err} ---")
                        traceback.print_exc()
                # ***** END ADD GRAND TOTAL ROW *****
 
                # --- V11: Logic for Summary Rows (BUFFALO summary + blank) ---
                # MOVED: This block now executes *after* GT row logic, conditional only on summary_flag
                summary_flag = sheet_mapping_section.get("summary", False)
                # Only proceed if summary flag is true AND processing was generally successful AND we have header info
                if summary_flag and processing_successful and last_table_header_info:
                    buffalo_summary_row = write_pointer_row # Target the *first* available row
                    blank_summary_row = write_pointer_row + 1 # The second row remains blank

                    print(f"DEBUG: summary flag is True. Preparing BUFFALO summary for row {buffalo_summary_row}.")

                    try:
                        # --- Define summary_font_to_use first, based on styling config or a default ---
                        # This logic mirrors how grand_total_font is defined elsewhere, ensuring consistency
                        # and availability even if the grand_total_font variable was not set (e.g., if num_tables <= 1).
                        summary_font_to_use = invoice_utils.bold_font # Default fallback

                        if sheet_styling_config:
                            # Prefer a specific "summary_font" config if available, otherwise fallback to "header_font"
                            font_config_key = "summary_font"
                            if font_config_key not in sheet_styling_config:
                                font_config_key = "header_font" # Fallback to using header_font style for summaries

                            actual_font_config = sheet_styling_config.get(font_config_key)
                            if actual_font_config and isinstance(actual_font_config, dict):
                                try:
                                    summary_font_to_use = openpyxl.styles.Font(
                                        name=actual_font_config.get("name", "Calibri"), # Default font name
                                        size=actual_font_config.get("size", 11),       # Default font size
                                        bold=actual_font_config.get("bold", True),     # Summaries are often bold
                                        italic=actual_font_config.get("italic", False),
                                        # color=actual_font_config.get("color", "000000") # Optional: Add color if in config
                                    )
                                    print(f"DEBUG: summary_font_to_use created from config key '{font_config_key}'.")
                                except Exception as font_err:
                                    print(f"Warning: Error creating summary_font_to_use from config '{font_config_key}': {font_err}. Using default bold font.")
                            else:
                                # This message indicates that neither "summary_font" nor "header_font" (as a fallback) was found in styling.
                                print(f"DEBUG: No specific font configuration found for '{font_config_key}' (or its fallback 'header_font') in sheet_styling_config for summary rows. Using default bold_font.")
                        else:
                            print("DEBUG: No sheet_styling_config provided for summary_font_to_use. Using default bold_font.")
                        # --- End summary_font_to_use definition ---

                        # --- Calculate BUFFALO and COW Totals (Including Pallets) ---
                        buffalo_totals = {"PCS": 0, "SF": 0, "N.W (kgs)": 0, "G.W (kgs)": 0, "CBM": 0}
                        cow_totals = {"PCS": 0, "SF": 0, "N.W (kgs)": 0, "G.W (kgs)": 0, "CBM": 0}
                        buffalo_pallet_total = 0
                        cow_pallet_total = 0
                        headers_to_sum = list(buffalo_totals.keys())

                        for table_key in table_keys:
                            table_data = all_tables_data.get(str(table_key))
                            if not table_data or not isinstance(table_data, dict): continue

                            descriptions = table_data.get("description", [])
                            pallet_counts = table_data.get("pallet_count", []) # Get pallet counts
                            max_len = len(descriptions)

                            pcs_list = table_data.get("pcs", [])
                            sf_list = table_data.get("sqft", [])
                            nw_list = table_data.get("net", [])
                            gw_list = table_data.get("gross", [])
                            cbm_list = table_data.get("cbm", [])

                            for i in range(max_len):
                                desc_val = descriptions[i] if i < len(descriptions) else None
                                is_buffalo = desc_val and "BUFFALO" in str(desc_val).upper()

                                target_dict = buffalo_totals if is_buffalo else cow_totals
                                current_pallet_count = 0
                                try: # Calculate current row pallet count safely
                                     if i < len(pallet_counts):
                                         raw_count = pallet_counts[i]
                                         current_pallet_count = int(raw_count) if isinstance(raw_count, (int, float)) or (isinstance(raw_count, str) and raw_count.isdigit()) else 0
                                         current_pallet_count = max(0, current_pallet_count)
                                except (ValueError, TypeError, IndexError): pass

                                # Add pallet count to respective total
                                if is_buffalo: buffalo_pallet_total += current_pallet_count
                                else: cow_pallet_total += current_pallet_count

                                # Add values to the appropriate dictionary safely
                                try: target_dict["PCS"] += int(pcs_list[i]) if i < len(pcs_list) and str(pcs_list[i]).isdigit() else 0
                                except (ValueError, TypeError, IndexError): pass
                                try: target_dict["SF"] += float(sf_list[i]) if i < len(sf_list) and isinstance(sf_list[i], (int, float, str)) else 0.0
                                except (ValueError, TypeError, IndexError): pass
                                try: target_dict["N.W (kgs)"] += float(nw_list[i]) if i < len(nw_list) and isinstance(nw_list[i], (int, float, str)) else 0.0
                                except (ValueError, TypeError, IndexError): pass
                                try: target_dict["G.W (kgs)"] += float(gw_list[i]) if i < len(gw_list) and isinstance(gw_list[i], (int, float, str)) else 0.0
                                except (ValueError, TypeError, IndexError): pass
                                try: target_dict["CBM"] += float(cbm_list[i]) if i < len(cbm_list) and isinstance(cbm_list[i], (int, float, str)) else 0.0
                                except (ValueError, TypeError, IndexError): pass

                        # --- Get Styling and Column Info ---
                        print(f"DEBUG: Writing BUFFALO summary to row {buffalo_summary_row} with totals: {buffalo_totals}, Pallets: {buffalo_pallet_total}")
                        print(f"DEBUG: Writing COW summary to row {blank_summary_row} with totals: {cow_totals}, Pallets: {cow_pallet_total}")
                        invoice_utils.unmerge_row(worksheet, buffalo_summary_row, last_table_header_info['num_columns'])
                        invoice_utils.unmerge_row(worksheet, blank_summary_row, last_table_header_info['num_columns'])
                        # The problematic line 'summary_font = grand_total_font' is now replaced by using summary_font_to_use directly.
                        column_map_gt = last_table_header_info['column_map']
                        idx_to_header_map_gt = {v: k for k, v in column_map_gt.items()}
                        # Find Description column index ONCE
                        desc_col_idx = column_map_gt.get("Description")
                        label_col_idx = column_map_gt.get("PALLET\nNO.") or column_map_gt.get("P.O Nº") or column_map_gt.get("P.O N°") or column_map_gt.get("CUT.P.O.")
                        if not label_col_idx:
                            for c_idx_fb in range(1, last_table_header_info['num_columns'] + 1):
                                 hdr_fb = idx_to_header_map_gt.get(c_idx_fb, "").lower()
                                 if "pallet" in hdr_fb or "po" in hdr_fb or "item" in hdr_fb: label_col_idx = c_idx_fb; break
                            if not label_col_idx: label_col_idx = 2 # Absolute fallback if no suitable column found

                        # --- Write BUFFALO Summary Row ---
                        if label_col_idx:
                            label_cell = worksheet.cell(row=buffalo_summary_row, column=label_col_idx, value="TOTAL OF:")
                            label_cell.font = summary_font_to_use # Apply the defined summary font
                            invoice_utils._apply_cell_style(label_cell, idx_to_header_map_gt.get(label_col_idx), sheet_styling_config)
                            description_bufflalo = worksheet.cell(row=buffalo_summary_row, column=label_col_idx + 1)
                            description_bufflalo.font = summary_font_to_use # Apply the defined summary font
                            description_bufflalo.alignment = openpyxl.styles.Alignment(horizontal='left', vertical='center')
                            description_bufflalo.value = "BUFFALO"

                        for header_name, total_value in buffalo_totals.items():
                             col_idx = column_map_gt.get(header_name)
                             if col_idx:
                                 sum_cell = worksheet.cell(row=buffalo_summary_row, column=col_idx, value=total_value)
                                 sum_cell.font = summary_font_to_use # Apply the defined summary font
                                 invoice_utils._apply_cell_style(sum_cell, header_name, sheet_styling_config)
                                 if header_name == "PCS": sum_cell.number_format = invoice_utils.FORMAT_NUMBER_COMMA_SEPARATED1
                                 elif header_name == "CBM": sum_cell.number_format = '0.00'
                                 else: sum_cell.number_format = invoice_utils.FORMAT_NUMBER_COMMA_SEPARATED2
                        # Write BUFFALO pallet total
                        if desc_col_idx:
                             pallet_cell_buffalo = worksheet.cell(row=buffalo_summary_row, column=desc_col_idx)
                             pallet_cell_buffalo.value = f"{buffalo_pallet_total} PALLETS"
                             pallet_cell_buffalo.font = summary_font_to_use # Apply the defined summary font
                             invoice_utils._apply_cell_style(pallet_cell_buffalo, "Description", sheet_styling_config)
                        # Apply borders and font to all cells in the BUFFALO summary row
                        for c_idx_sum in range(1, last_table_header_info['num_columns'] + 1):
                             cell = worksheet.cell(row=buffalo_summary_row, column=c_idx_sum)
                            #  cell.border = invoice_utils.thin_border # Uncomment to apply border
                             if cell.value is not None: cell.font = summary_font_to_use # Ensure font is applied

                        # --- Write COW Summary Row ---
                        if label_col_idx:
                            label_cell_cow = worksheet.cell(row=blank_summary_row, column=label_col_idx, value="TOTAL OF:") # Use blank_summary_row
                            label_cell_cow.font = summary_font_to_use # Apply the defined summary font
                            invoice_utils._apply_cell_style(label_cell_cow, idx_to_header_map_gt.get(label_col_idx), sheet_styling_config)
                            description_cow = worksheet.cell(row=blank_summary_row, column=label_col_idx + 1)
                            description_cow.font = summary_font_to_use # Apply the defined summary font
                            description_cow.alignment = openpyxl.styles.Alignment(horizontal='left', vertical='center')
                            description_cow.value = "LEATHER" # Changed from "COW" to "LEATHER" as per original code

                        for header_name, total_value in cow_totals.items(): # Use cow_totals
                             col_idx = column_map_gt.get(header_name)
                             if col_idx:
                                 sum_cell_cow = worksheet.cell(row=blank_summary_row, column=col_idx, value=total_value)
                                 sum_cell_cow.font = summary_font_to_use # Apply the defined summary font
                                 invoice_utils._apply_cell_style(sum_cell_cow, header_name, sheet_styling_config)
                                 if header_name == "PCS": sum_cell_cow.number_format = invoice_utils.FORMAT_NUMBER_COMMA_SEPARATED1
                                 elif header_name == "CBM": sum_cell_cow.number_format = '0.00'
                                 else: sum_cell_cow.number_format = invoice_utils.FORMAT_NUMBER_COMMA_SEPARATED2
                        # Write COW pallet total
                        if desc_col_idx:
                            pallet_cell_cow = worksheet.cell(row=blank_summary_row, column=desc_col_idx)
                            pallet_cell_cow.value = f"{cow_pallet_total} PALLETS"
                            pallet_cell_cow.font = summary_font_to_use # Apply the defined summary font
                            invoice_utils._apply_cell_style(pallet_cell_cow, "Description", sheet_styling_config)
                        # Apply borders and font to all cells in the COW summary row
                        for c_idx_sum in range(1, last_table_header_info['num_columns'] + 1):
                             cell = worksheet.cell(row=blank_summary_row, column=c_idx_sum)
                            #  cell.border = invoice_utils.thin_border # Uncomment to apply border
                             if cell.value is not None: cell.font = summary_font_to_use # Ensure font is applied

                        # Advance pointer past the two summary rows (one written, one blank)
                        write_pointer_row += 2
                        print(f"DEBUG: Finished BUFFALO & COW summaries. Next pointer is now {write_pointer_row}.")

                    except Exception as summary_err:
                        print(f"Warning: Failed processing BUFFALO summary rows: {summary_err}")
                        traceback.print_exc()
                        # Decide if pointer should still advance? Maybe not if error is severe.
                        # For now, let's advance it to avoid potential overlaps later.
                        write_pointer_row += 2 # Advance pointer even on error
                # --- End Summary Rows Logic ---

                # --- START FOB-Specific Hardcoded Replacements (Multi-Table) ---
                if args.fob:
                    # Apply to the current multi-table sheet if --fob is active
                    print(f":::::::::::::::::::: Performing FOB-specific hardcoded replacements for multi-table sheet '{sheet_name}'...")
                    fob_specific_replacement_rules = [
                        {
                            "find": "DAP",          # Text to find
                            "replace": "FOB",       # Text to replace with
                            "case_sensitive": False # Match regardless of case
                        },
                        {
                            "find": "FCA",          # Text to find
                            "replace": "FOB",       # Text to replace with
                            "case_sensitive": False # Match regardless of case
                        },
                        {
                            "find": "BAVET, SVAY RIENG",    # Text to find
                            "replace": "BAVET",     # Text to replace with
                            "case_sensitive": True, # Match regardless of case
                            "exact_cell_match": True
                        },
                        {
                            "find": "BINH PHUOC",    # Text to find
                            "replace": "BAVET",     # Text to replace with
                            "case_sensitive": True, # Match regardless of case
                            "exact_cell_match": True
                        }
                        # Add more FOB-specific rules here if needed
                    ]
                    # Call the function designed for find/replace rules
                    # Apply ONLY to the current sheet being processed
                    text_replace_utils.find_and_replace_in_workbook(
                        workbook=workbook,
                        replacement_rules=fob_specific_replacement_rules,
                        target_sheets=[sheet_name] # Apply only to the current sheet
                    )
                    print(f":::::::::::::::::::: Finished FOB-specific hardcoded replacements for multi-table sheet '{sheet_name}'.")
                # --- END FOB-Specific Hardcoded Replacements (Multi-Table) ---

                # --- Apply Column Widths AFTER loop using the last header info ---
                if processing_successful and last_table_header_info:
                    print(f"Applying column widths for multi-table sheet '{sheet_name}'...")
                    invoice_utils.apply_column_widths(
                        worksheet,
                        sheet_styling_config,
                        last_table_header_info.get('column_map')
                    )
                # --- End Apply Column Widths ---
 
                # --- Final Spacer Rows --- 
                # V11: No insert needed, just advance pointer if required
                if final_row_spacing > 0 and num_tables > 0 and processing_successful:
                    final_spacer_start_row = write_pointer_row
                    try:
                        print(f"Config requests final spacing ({final_row_spacing}). Advancing pointer from {final_spacer_start_row}.")
                        # worksheet.insert_rows(final_spacer_start_row, amount=final_row_spacing) # REMOVED INSERT
                        # Optionally clear/style these rows
                        write_pointer_row += final_row_spacing
                        print(f"Pointer advanced for final spacing. Final pointer: {write_pointer_row}")
                    except Exception as final_spacer_err: 
                        print(f"Warning: Error during final spacing logic: {final_spacer_err}")
 
            # ================================================================
            # --- Handle Single Table / Aggregation Case ---
            # ================================================================
            else:
                print(f"Processing sheet '{sheet_name}' as single table/aggregation.")
                header_info = None; footer_info = None

                start_row = sheet_mapping_section.get('start_row');
                header_to_write = sheet_mapping_section.get('header_to_write');
                header_merge_rules = sheet_mapping_section.get('header_merge_rules')
                if not start_row or not header_to_write: print(f"Error: Config for '{sheet_name}' missing 'start_row' or 'header_to_write'. Skipping."); processing_successful = False; continue

                print(f"Writing header at row {start_row}...");
                header_info = invoice_utils.write_header(
                    worksheet, start_row, header_to_write, header_merge_rules, sheet_styling_config
                )
                if not header_info: print(f"Error: Failed to write header for '{sheet_name}'. Skipping."); processing_successful = False; continue
                print(f"DEBUG: Header Info for '{sheet_name}':")
                print(f"  - Column Map: {header_info.get('column_map')}")
                print(f"Header written successfully.")

                # --- Find Footer --- (Existing logic)
                if global_footer_rules and global_footer_rules.get('marker_text'):
                     print("Finding footer marker...");
                     footer_info = invoice_utils.find_footer(worksheet, global_footer_rules)
                     if not footer_info: print(f"Warning: Footer marker '{global_footer_rules.get('marker_text')}' not found.")
                else: print(f"Footer marker found: {footer_info}")

                # --- Get Data Source ---
                data_to_fill = None; data_source_type = None
                print(f"DEBUG: Retrieving data source for '{sheet_name}' using indicator: '{data_source_indicator}'")

                # --- Custom Flag Logic ---
                if args.custom and data_source_indicator == 'aggregation':
                    print(f"DEBUG: --custom flag active. Attempting to use 'custom_aggregation_results' for sheet '{sheet_name}'.")
                    data_to_fill = invoice_data.get('custom_aggregation_results')
                    data_source_type = 'aggregation' # Type remains aggregation
                    if data_to_fill is not None:
                        print("DEBUG: Successfully retrieved 'custom_aggregation_results'.")
                    else:
                        print("Warning: 'custom_aggregation_results' key not found in data, despite --custom flag.")
                # --- End Custom Flag Logic ---

                # --- Standard Data Source Selection (if not overridden by --custom) ---
                if data_to_fill is None: # Only proceed if custom logic didn't already set it
                    if data_source_indicator == 'fob_aggregation': # Check for FOB first
                        data_to_fill = invoice_data.get('final_fob_compounded_result'); data_source_type = 'fob_aggregation';
                        print("DEBUG: Using 'final_fob_compounded_result' data (FOB mode).")
                        
                        # Add mapping for combined_description to Description column
                        if isinstance(data_to_fill, dict):
                            for map_key, map_rule in sheet_inner_mapping_rules_dict.items():
                                if isinstance(map_rule, dict):
                                    header = map_rule.get('header')
                                    if header in ['Description', 'DESCRIPTION OF GOODS', 'Description of Goods']:
                                        # Override any existing mapping to use combined_description
                                        map_rule['value_key'] = 'combined_description'
                                        print(f"DEBUG: FOB Mode - Mapped combined_description to {header}")
                    elif data_source_indicator == 'aggregation': # Standard aggregation
                        data_to_fill = invoice_data.get('standard_aggregation_results'); data_source_type = 'aggregation';
                        print("DEBUG: Using 'standard_aggregation_results' data.")
                    elif 'processed_tables_data' in invoice_data and data_source_indicator in invoice_data.get('processed_tables_data', {}):
                        data_to_fill = invoice_data['processed_tables_data'].get(data_source_indicator); data_source_type = 'processed_tables';
                        print(f"DEBUG: Using 'processed_tables_data' key '{data_source_indicator}' (Likely for multi-table).")
                    else:
                        print(f"DEBUG: Data source indicator '{data_source_indicator}' not found in expected locations ('final_fob_compounded_result', 'standard_aggregation_results', 'custom_aggregation_results' or 'processed_tables_data').")
                # --- End Standard Data Source Selection ---

                if data_to_fill is None: print(f"Warning: Data source '{data_source_indicator}' unknown or data empty. Skipping fill."); continue

                # *** Explicitly re-check header_info before the call ***
                # header_info = written_header_info # Re-assign just in case (Use the variable name from the write_header call if different)
                # Let's assume the variable was indeed header_info
                if header_info is None:
                    print(f"FATAL DEBUG: header_info is unexpectedly None right before calling fill_invoice_data for {sheet_name}")

                if header_info is not None and header_info.get('column_map'):
                    print(f"DEBUG: Calling fill_invoice_data for single table '{sheet_name}'")
                    print(f"DEBUG: header_info passed to fill_invoice_data: {header_info}") # <--- SHOWS None HERE
                    # *** Get aggregated FOB values ---
                    total_sqft = data_to_fill.get('total_sqft', 0)
                    total_amount = data_to_fill.get('total_amount', 0)
                    combined_po = data_to_fill.get('combined_po')
                    combined_item = data_to_fill.get('combined_item')
                    combined_description = data_to_fill.get('combined_description')

                    # *** Map headers to data source keys and find target columns ---
                    fob_data_keys = {
                        'P.O N °': 'combined_po', 'P.O Nº': 'combined_po',
                        'ITEM NO': 'combined_item', 'ITEM Nº': 'combined_item',
                        'Quantity ( SF )': 'total_sqft', 'Quantity(SF)': 'total_sqft',
                        'Amount ( USD )': 'total_amount', 'Total value(USD)': 'total_amount',
                        'Description': 'combined_description', 'DESCRIPTION OF GOODS': 'combined_description',
                        'Description of Goods': 'combined_description', "Description Of Goods": 'combined_item'
                    }

                    # ***** MODIFIED CALL for single table (Removed starting_pallet_order, adjusted return) *****
                    fill_success, next_row_after_footer, dog, _, table_pallets = invoice_utils.fill_invoice_data(
                        worksheet=worksheet,
                        sheet_name=sheet_name,
                        sheet_config=sheet_mapping_section, # Pass current sheet's config
                        all_sheet_configs=data_mapping_config, # <--- Pass the full config map
                        data_source=data_to_fill,
                        data_source_type=data_source_type,
                        header_info=header_info,
                        mapping_rules=sheet_inner_mapping_rules_dict,
                        sheet_styling_config=sheet_styling_config,
                        add_blank_after_header=add_blank_after_hdr_flag,
                        static_content_after_header=static_content_after_hdr_dict,
                        add_blank_before_footer=add_blank_before_ftr_flag,
                        static_content_before_footer=static_content_before_ftr_dict,
                        merge_rules_after_header=merge_rules_after_hdr,
                        merge_rules_before_footer=merge_rules_before_ftr,
                        merge_rules_footer=merge_rules_footer,
                        footer_info=footer_info, max_rows_to_fill=None,
                        grand_total_pallets=final_grand_total_pallets,
                        custom_flag=args.custom,
                        data_cell_merging_rules=data_cell_merging_rules
                    )
                    # ***** END MODIFIED CALL *****
                    
                    # Initialize sheet_grand_totals dictionary
                    sheet_grand_totals = {
                        "grand_total_nett_weight": 0.0,
                        "grand_total_gross_weight": 0.0,
                        "grand_total_cbm": 0.0,  # <-- STEP 1: ADDED NEW KEY FOR CBM
                        # Add other totals needed for rows_after_footer here
                    }

                    # This block calculates totals if the sheet is a multi-table sheet (e.g. Packing List)
                    # and these totals are then used by write_configured_rows.
                    # The condition `processing_successful and all_tables_data` applies.
                    # `all_tables_data` is typically `invoice_data.get('processed_tables_data', {})`
                    # `table_keys` are the sorted keys from `all_tables_data`
                    if processing_successful and all_tables_data: # This condition might be more relevant for multi-table sheets
                        print("Calculating totals for rows_after_footer...")
                        for table_key in table_keys: # table_keys would be derived from all_tables_data
                            table_data = all_tables_data.get(str(table_key))
                            if table_data and isinstance(table_data, dict):
                                try:
                                    # Summing 'net' values - ensure 'net' exists and contains numbers
                                    nett_weights = table_data.get("net", [])
                                    if isinstance(nett_weights, list):
                                        sheet_grand_totals["grand_total_nett_weight"] += sum(
                                            float(nw) for nw in nett_weights 
                                            if isinstance(nw, (int, float, str)) and str(nw).strip() # Check if convertible
                                        )
                                except (ValueError, TypeError) as e:
                                    print(f"Warning: Error summing nett weights for table {table_key}: {e}")
                                try:
                                    # Summing 'gross' values - ensure 'gross' exists and contains numbers
                                    gross_weights = table_data.get("gross", [])
                                    if isinstance(gross_weights, list):
                                        sheet_grand_totals["grand_total_gross_weight"] += sum(
                                            float(gw) for gw in gross_weights 
                                            if isinstance(gw, (int, float, str)) and str(gw).strip() # Check if convertible
                                        )
                                except (ValueError, TypeError) as e:
                                    print(f"Warning: Error summing gross weights for table {table_key}: {e}")
                                
                                # --- STEP 2: ADDED CALCULATION LOGIC FOR CBM ---
                                try:
                                    # Summing 'cbm' values - ensure 'cbm' exists and contains numbers
                                    cbm_values = table_data.get("cbm", []) # Assuming 'cbm' is the key in your data
                                    if isinstance(cbm_values, list):
                                        sheet_grand_totals["grand_total_cbm"] += sum(
                                            float(cbm_val) for cbm_val in cbm_values 
                                            if isinstance(cbm_val, (int, float, str)) and str(cbm_val).strip() # Check if convertible
                                        )
                                except (ValueError, TypeError) as e:
                                    print(f"Warning: Error summing CBM for table {table_key}: {e}")
                                # --- END STEP 2 ---

                        print(f"Calculated Sheet Grand Totals: {sheet_grand_totals}")
                        
                        # This part calls write_configured_rows. 
                        # Ensure last_table_header_info is relevant or use header_info for single table sheets.
                        # For single table sheets, header_info should be used.
                        # The write_pointer_row also needs to be correctly set. For single table, it would be next_row_after_footer.
                        
                        current_header_info_for_rows_after_footer = header_info # For single table sheets
                        current_write_pointer_for_rows_after_footer = next_row_after_footer

                        # If this logic is inside the "processed_tables_multi" part of your script, 
                        # then last_table_header_info and write_pointer_row (as updated by the multi-table loop) would be correct.
                        # The provided snippet seems to be from the single-table processing path due to `next_row_after_footer`.

                        if processing_successful and current_header_info_for_rows_after_footer:
                            rows_after_footer_enabled = sheet_mapping_section.get("rows_after_footer_enabled", False)
                            rows_after_footer_config = sheet_mapping_section.get("rows_after_footer", [])
                            
                            if rows_after_footer_enabled and isinstance(rows_after_footer_config, list) and len(rows_after_footer_config) > 0:
                                num_extra_rows = len(rows_after_footer_config)
                                
                                # For single-table sheets, insert rows if they are not already accounted for.
                                # This depends on whether fill_invoice_data already inserted space for these.
                                # Typically, fill_invoice_data inserts rows up to its own footer.
                                # So, rows *after* that footer need to be inserted if not already planned.
                                # However, write_configured_rows assumes rows are ALREADY INSERTED.
                                # So, you might need to insert rows *before* calling write_configured_rows.

                                print(f"DEBUG: Preparing to write {num_extra_rows} rows after footer, starting at potential row {current_write_pointer_for_rows_after_footer}")
                                print(f"DEBUG: Using header_info: {current_header_info_for_rows_after_footer.get('num_columns')} columns")
                                print(f"DEBUG: Calculated totals for these rows: {sheet_grand_totals}")
                                if num_extra_rows > 0: 
                                    print(f"DEBUG: Inserting {num_extra_rows} rows at {current_write_pointer_for_rows_after_footer} for 'rows_after_footer'.")
                                    worksheet.insert_rows(current_write_pointer_for_rows_after_footer, amount=num_extra_rows)
                                    # Unmerge the newly inserted block to prevent styling/merge issues
                                    invoice_utils.safe_unmerge_block(
                                        worksheet, 
                                        current_write_pointer_for_rows_after_footer, 
                                        current_write_pointer_for_rows_after_footer + num_extra_rows - 1, 
                                        current_header_info_for_rows_after_footer['num_columns']
                                    )
                                    print(f"DEBUG: Inserted and unmerged {num_extra_rows} rows.")
                                else:
                                    print("DEBUG: No extra rows configured in 'rows_after_footer', no insertion needed.")

                                invoice_utils.write_configured_rows( 
                                    worksheet=worksheet,
                                    start_row_index=current_write_pointer_for_rows_after_footer, # This should be the row where these new lines start
                                    num_columns=current_header_info_for_rows_after_footer['num_columns'],
                                    rows_config_list=rows_after_footer_config,
                                    calculated_totals=sheet_grand_totals, 
                                    default_style_config=sheet_styling_config
                                )
                                next_row_after_footer = next_row_after_footer + 2

                                # Update the pointer if you were managing it manually for subsequent operations on this sheet
                                # current_write_pointer_for_rows_after_footer += num_extra_rows 
                                # next_row_after_footer = current_write_pointer_for_rows_after_footer # If this is the very end of sheet processing

                                print(f"DEBUG: Finished writing rows after footer. The next available row would conceptually be after these {num_extra_rows} rows.")
                    
                    


                    if fill_success:
                        print(f"Successfully filled table data/footer for sheet '{sheet_name}'.")

                        # --- Apply Column Widths AFTER filling ---
                        print(f"Applying column widths for sheet '{sheet_name}'...")
                        invoice_utils.apply_column_widths(
                            worksheet,
                            sheet_styling_config,
                            header_info.get('column_map')
                        )
                        # --- End Apply Column Widths ---

                        # --- Final Spacer Rows --- (Existing logic)
                        if final_row_spacing >= 1:
                            final_spacer_start_row = next_row_after_footer
                            try:
                                print(f"Config requests final spacing ({final_row_spacing}). Adding blank row(s) at {final_spacer_start_row}.")
                                worksheet.insert_rows(final_spacer_start_row, amount=final_row_spacing)
                            except Exception as final_spacer_err: print(f"Warning: Failed to insert final spacer rows: {final_spacer_err}")

                        # --- Fill Summary Fields --- (Existing logic)
                        print("Attempting to fill summary fields...")
                        summary_data_source = invoice_data.get('final_fob_compounded_result', {})
                        if not summary_data_source: print("Warning: 'final_fob_compounded_result' not found.")
                        elif not sheet_inner_mapping_rules_dict: print("Warning: 'mappings' dict not found for summary fields.")
                        else:
                            summary_fields_found = 0
                            for map_key, map_rule in sheet_inner_mapping_rules_dict.items():
                                if isinstance(map_rule, dict) and 'marker' in map_rule:
                                    marker_text = map_rule['marker']; summary_value = summary_data_source.get(map_key)
                                    if not marker_text: print(f"Warning: Marker text missing for summary key '{map_key}'."); continue
                                    if summary_value is None: print(f"Warning: Summary value for key '{map_key}' not found."); continue
                                    target_cell = invoice_utils.find_cell_by_marker(worksheet, marker_text)
                                    if target_cell:
                                        summary_fields_found += 1
                                        try: # Format as number if possible
                                            summary_value_num = float(summary_value);
                                            target_cell.value = summary_value_num; target_cell.number_format = '#,##0.00' # Default summary format
                                            print(f"Wrote summary {summary_value_num:.2f} to {target_cell.coordinate} ('{marker_text}')")
                                        except (ValueError, TypeError): # Fallback to string
                                            target_cell.value = str(summary_value)
                                            print(f"Wrote summary '{summary_value}' as string to {target_cell.coordinate} ('{marker_text}')")
                            if summary_fields_found == 0: print("No summary field markers found/matched.")
                            else: print(f"Processed {summary_fields_found} summary field(s).")

                        # --- Define and Apply Data-Driven Replacements HERE (Inside single-table loop) ---
                        print("Performing data-driven replacements for single-table sheet...")
                        # Define rules here: placeholder text, path to data within invoice_data, target sheets
                        # NOTE: Target sheets here might be redundant if we only apply to the *current* sheet
                        # Consider refining rules if replacement should only apply to the sheet being processed.
                        data_driven_replacement_rules = [
                            {
                                "find": "JFINV",
                                "data_path": ["processed_tables_data", "1", "inv_no", 0], # Path within invoice_data
                                "target_sheets": ["Invoice", "Contract", "Packing list"], # Apply ONLY to current sheet
                                "case_sensitive": True
                            },
                            {
                                "find": "JFTIME",
                                "data_path": ["processed_tables_data", "1", "inv_date", 0], # Corrected path based on JF.json analysis
                                "target_sheets": ["Invoice", "Contract", "Packing list"], # Apply ONLY to current sheet
                                "case_sensitive": True,
                                "is_date": True  # Mark this as a date field for special handling
                            },
                            {
                                "find": "JFREF",
                                # IMPORTANT: Determine the CORRECT key for reference from JF.json
                                # Using 'reference_code' as a placeholder - CHANGE IF WRONG
                                "data_path": ["processed_tables_data", "1", "inv_ref", 0],
                                "target_sheets": ["Invoice", "Contract", "Packing list"], # Apply ONLY to current sheet
                                "case_sensitive": True
                            },
                            # Add more rules as needed
                        ]
                        # Call the utility function
                        # Add the :::::::::::::::::::: debug print here
                        proc_tables_rep = invoice_data.get('processed_tables_data', {})
                        table_1_data_rep = proc_tables_rep.get('1', {})

                        text_replace_utils.process_data_driven_replacements(
                            workbook,
                            invoice_data,
                            data_driven_replacement_rules
                        )
                        # --- End Data-Driven Replacements for single-table ---

                        # --- START FOB-Specific Hardcoded Replacements ---
                        if args.fob:
                            # Apply to ALL sheets when --fob is active (but only runs when processing that specific sheet)
                            print(f":::::::::::::::::::: Performing FOB-specific hardcoded replacements for sheet '{sheet_name}'...")
                            fob_specific_replacement_rules = [
                                {
                                    "find": "DAP",          # Text to find
                                    "replace": "FOB",       # Text to replace with
                                    "case_sensitive": False # Match regardless of case
                                },
                                {
                                    "find": "FCA",          # Text to find
                                    "replace": "FOB",       # Text to replace with
                                    "case_sensitive": False # Match regardless of case
                                },
                                {
                                    "find": "BINH PHUOC",    # Text to find
                                    "replace": "BAVET",     # Text to replace with
                                    "case_sensitive": False, # Match regardless of case
                                    "exact_cell_match": True
                                },
                                {
                                    "find": "BAVET, SVAY RIENG",    # Text to find
                                    "replace": "BAVET",     # Text to replace with
                                    "case_sensitive": True, # Match regardless of case
                                    "exact_cell_match": True
                                }
                                # Add more FOB-specific rules here if needed
                            ]
                            # Call the function designed for find/replace rules
                            # The function will apply the rules only to the sheet specified in target_sheets
                            text_replace_utils.find_and_replace_in_workbook(
                                workbook=workbook,
                                replacement_rules=fob_specific_replacement_rules,
                                target_sheets=[sheet_name] # Apply only to the current sheet being processed
                            )
                            print(f":::::::::::::::::::: Finished FOB-specific hardcoded replacements for sheet '{sheet_name}'.")
                        # --- END FOB-Specific Hardcoded Replacements ---

                    else: print(f"Failed to fill table data/footer for sheet '{sheet_name}'."); processing_successful = False
                else:
                    print(f"Error: Cannot fill data for '{sheet_name}' because header_info or column_map is missing/invalid.")
                    processing_successful = False; continue # Skip filling if header info is bad

        # --- Restore Original Merges AFTER processing all sheets using merge_utils ---
        merge_utils.find_and_restore_merges_heuristic(workbook, original_merges, sheets_to_process)

        # 5. Save the final workbook
        print("\n--------------------------------")
        if processing_successful:
            print("5. Saving final workbook...")
            workbook.save(output_path); print(f"--- Workbook saved successfully: '{output_path}' ---")
        else:
            print("--- Processing completed with errors. Saving workbook (may be incomplete). ---")
            try:
                # Corrected the closing quote below
                workbook.save(output_path); print(f"--- Incomplete workbook saved to: '{output_path}' ---")
            except Exception as save_err:
                print(f"--- CRITICAL ERROR: Failed to save incomplete workbook: {save_err} ---")

    except Exception as e:
        print(f"\n--- UNHANDLED ERROR during workbook processing: {e} ---"); traceback.print_exc()
        if workbook and output_path: # Try to save error state
             try:
                 error_filename = output_path.stem + "_ERROR" + output_path.suffix; error_path = output_path.with_name(error_filename)
                 print(f"Attempting to save workbook state to {error_path}..."); workbook.save(error_path); print("Workbook state saved.")
             except Exception as final_save_err: print(f"--- Could not save workbook state after error: {final_save_err} ---")
    finally:
        if workbook:
            try: workbook.close(); print("Workbook closed.")
            except Exception: pass

    print("\n--- Invoice Generation Finished ---")

# --- Run Main ---
if __name__ == "__main__":
    main()