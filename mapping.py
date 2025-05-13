import sys
import io
import openpyxl
from openpyxl import load_workbook
import json
import argparse


sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
def load_mappings(attribute_json_path):
    try:
        with open(attribute_json_path, 'r', encoding='utf-8') as f:
            mappings = json.load(f)
        return mappings
    except Exception as e:
        print(f"Error loading attribute JSON: {str(e)}")
        raise

def find_matching_sheet(source_wb, mappings, sheet_type):
    for sheet_name in source_wb.sheetnames:
        variations = mappings["Trường data"].get(sheet_type, {}).get(sheet_type, [])
        if any(sheet_name.lower() == variation.lower() for variation in variations):
            return source_wb[sheet_name]
    return None

def map_fields_to_columns(source_sheet, mappings, sheet_type):
    header_row = next(source_sheet.iter_rows(min_row=1, max_row=1, values_only=True))
    fields = mappings["Trường data"].get(sheet_type, {}).keys()
    field_names = list(fields)[1:]  # Remove the sheet name field
    
    field_columns = {}
    for idx, header in enumerate(header_row, start=1):
        if header:
            header_lower = str(header).lower()
            for field in mappings["Trường data"].get(sheet_type, {}):
                for variation in mappings["Trường data"][sheet_type][field]:
                    if header_lower == variation.lower():
                        field_columns[field] = idx
                        break  # Break after first match is found

    if len(field_columns) < len(field_names):
        missing = [field for field in field_names if field not in field_columns]
        raise ValueError(f"Required columns not found in {sheet_type}: {', '.join(missing)}")
    
    return field_columns, field_names

def map_target_columns(target_sheet, field_names):
    target_header_row = next(target_sheet.iter_rows(min_row=1, max_row=1, values_only=True))
    field_columns_targeted = {}
    for idx, header in enumerate(target_header_row, start=1):
        if header:
            header_str = str(header).strip()
            if header_str in field_names:
                field_columns_targeted[header_str] = idx
    return field_columns_targeted

def copy_data_between_sheets(source_sheet, target_sheet, field_columns, target_columns, field_names):
    for row_idx in range(2, source_sheet.max_row + 1):  # Skip header row
        row_data = [source_sheet.cell(row=row_idx, column=field_columns[field]).value for field in field_names]
        for field, col_idx in target_columns.items():
            target_sheet.cell(row=row_idx, column=col_idx).value = row_data[field_names.index(field)]

def process_sheet(source_wb, target_wb, mappings, sheet_type):
    try:
        source_sheet = find_matching_sheet(source_wb, mappings, sheet_type)
        if not source_sheet:
            print(f"No matching '{sheet_type}' sheet found in source file - skipping")
            return
        
        target_sheet = target_wb[sheet_type]
        field_columns, field_names = map_fields_to_columns(source_sheet, mappings, sheet_type)
        target_columns = map_target_columns(target_sheet, field_names)
        copy_data_between_sheets(source_sheet, target_sheet, field_columns, target_columns, field_names)
        print(f"Successfully processed {sheet_type} sheet")
    except Exception as e:
        print(f"Error processing {sheet_type} sheet: {str(e)}")
        raise

def mapping(filepath, formated, attribute_json_path):
    try:
        mappings = load_mappings(attribute_json_path)
        file_khach_hang = load_workbook(filepath)
        file_format = load_workbook(formated)

        for sheet_type in mappings["Trường data"]:
            process_sheet(file_khach_hang, file_format, mappings, sheet_type)
        
        # Save and return the formatted file path
        file_format.save(formated)
        print(f"Successfully copied data from {filepath} to {formated}")
        return formated

    except Exception as e:
        print(f"Error: {str(e)}")
        raise

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Map data from source Excel to formatted Excel based on JSON attributes")
    parser.add_argument("filepath", help="Path to the source Excel file")
    parser.add_argument("formated", help="Path to the formatted Excel file")
    parser.add_argument("atribute_json", help="Path to the attribute JSON file")
    args = parser.parse_args()
    
    # Run the mapping and get the formatted file path
    result_file = mapping(args.filepath, args.formated, args.atribute_json)
    print(f"Formatted file saved at: {result_file}")