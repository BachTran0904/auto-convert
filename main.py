import sys, re
import io
import openpyxl
from openpyxl import load_workbook
from openpyxl.styles import Font, Alignment, PatternFill
from openpyxl.utils import get_column_letter
import json

import json
from openpyxl import load_workbook

def fill_page_hang_hoa(filepath, formated, attribute_json_path):
    try:
        # Load attribute mappings
        with open(attribute_json_path, 'r', encoding='utf-8') as f:
            mappings = json.load(f)
        
        # Load the workbooks
        file_khach_hang = load_workbook(filepath)
        file_format = load_workbook(formated)

        # Find the correct sheet in source file
        source_sheet = None
        for sheet_name in file_khach_hang.sheetnames:
            for variation in mappings["Trường data"]["Hàng hóa"]["Hàng hóa"]:
                if sheet_name.lower() == variation.lower():
                    source_sheet = file_khach_hang[sheet_name]
                    break
            if source_sheet:
                break
        
        if not source_sheet:
            raise ValueError("No matching 'Hàng hóa' sheet found in source file")

        
        # Get target sheet
        trang_hang_hoa_format = file_format['Hàng hóa']
        
        # Find columns in source sheet using attribute mappings
        header_row = next(source_sheet.iter_rows(min_row=1, max_row=1, values_only=True))
        
        ma_sp_col = None
        ten_sp_col = None
        UOM1_col = None
        
        for idx, header in enumerate(header_row, start=1):
            if header:
                header_lower = str(header).lower()
                # Check for Mã hàng variations
                for variation in mappings["Trường data"]["Hàng hóa"]["Mã hàng"]:
                    if header_lower == variation.lower():
                        ma_sp_col = idx
                        break
                # Check for Tên hàng variations
                for variation in mappings["Trường data"]["Hàng hóa"]["Tên hàng"]:
                    if header_lower == variation.lower():
                        ten_sp_col = idx
                        break
                # Check for UOM variations
                for variation in mappings["Trường data"]["Hàng hóa"]["UOM"]:
                    if header_lower == variation.lower():
                        UOM1_col = idx
                        break
        
        if not all([ma_sp_col, ten_sp_col, UOM1_col]):
            missing = []
            if not ma_sp_col: missing.append("Mã hàng")
            if not ten_sp_col: missing.append("Tên hàng")
            if not UOM1_col: missing.append("UOM")
            raise ValueError(f"Required columns not found: {', '.join(missing)}")

        # Find columns in target sheet
        target_header_row = next(trang_hang_hoa_format.iter_rows(min_row=1, max_row=1, values_only=True))
        
        ma_hang_col = None
        ten_hang_col = None
        UOM_col = None
        
        for idx, header in enumerate(target_header_row, start=1):
            if header:
                header_str = str(header).strip()
                if header_str == "Mã hàng":
                    ma_hang_col = idx
                elif header_str == "Tên hàng":
                    ten_hang_col = idx
                elif header_str == "UOM":
                    UOM_col = idx
        
        if not all([ma_hang_col, ten_hang_col, UOM_col]):
            raise ValueError("Required columns not found in target template")

        # Copy data from source to target
        for row_idx in range(1, source_sheet.max_row):
            ma_sp_value = source_sheet.cell(row=row_idx, column=ma_sp_col).value
            ten_sp_value = source_sheet.cell(row=row_idx, column=ten_sp_col).value
            UOM1_value = source_sheet.cell(row=row_idx, column=UOM1_col).value

            # Write to target sheet (same row index)
            trang_hang_hoa_format.cell(row=row_idx, column=ma_hang_col, value=ma_sp_value)
            trang_hang_hoa_format.cell(row=row_idx, column=ten_hang_col, value=ten_sp_value)
            trang_hang_hoa_format.cell(row=row_idx, column=UOM_col, value=UOM1_value)

        # Save the workbook
        file_format.save(formated)
        print(f"Successfully copied data from {filepath} to {formated}")
        
    except Exception as e:
        print(f"Error: {str(e)}")
        raise
        

        
#Call
filepath = 'Raw.xlsx'
formated = 'Form.xlsx'
atribute_json = 'atribute.json'
fill_page_hang_hoa(filepath, formated, atribute_json)


