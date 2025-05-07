import sys, re
import io
import openpyxl
from openpyxl import load_workbook
from openpyxl.styles import Font, Alignment, PatternFill
from openpyxl.utils import get_column_letter
import json

pattern = re.compile(r'hang\s+hoa', re.IGNORECASE)

def fill_page_hang_hoa(filepath, formated):
    try:
        # Load the workbook
        file_khach_hang = load_workbook(filepath)
        file_format = load_workbook(formated)
        trang_hang_hoa_format = file_format['Hàng hóa']

        if pattern.search(normalize("Hàng hóa")):
            trang_hang_hoa = file_khach_hang['HÀNG HÓA']

        # Find "Mã SP" column in raw sheet
        ma_sp_col = None
        ten_sp_col = None
        UOM1_col = None
        first_flag = None
        second_flag = None
        third_flag = None
        for cell in trang_hang_hoa[1]:  # Check first row for header
            if cell.value == "Mã SP":
                ma_sp_col = cell.column
                first_flag = True
            if cell.value == "Tên SP":
                ten_sp_col = cell.column
                second_flag = True
            if cell.value == "UOM1":
                UOM1_col = cell.column
                third_flag = True
            if ((first_flag == True) & (second_flag == True) & (third_flag == True)):
                break
        


        # Find or create "Mã Hàng" column in target sheet
        ma_hang_col = None
        for cell in trang_hang_hoa_format[1]:  # Check first row for header
            if cell.value == "Mã hàng":
                ma_hang_col = cell.column
                break     
        ten_hang_col = None
        for cell in trang_hang_hoa_format[1]:  # Check first row for header
            if cell.value == "Tên hàng":
                ten_hang_col = cell.column
                break
        UOM_col = None
        for cell in trang_hang_hoa_format[1]:  # Check first row for header
            if cell.value == "UOM":
                UOM_col = cell.column
                break

        # Copy data from "Mã SP" to "Mã Hàng"
        for row_idx in range(2, file_khach_hang.max_row + 1):  # Start from row 2 to skip header
            ma_sp_value = trang_hang_hoa.cell(row=row_idx, column=ma_sp_col).value
            ten_sp_value = trang_hang_hoa.cell(row=row_idx, column=ten_sp_col).value
            UOM1_value = trang_hang_hoa.cell(row=row_idx, column=UOM1_col).value

            # Write to target sheet (same row index)
            trang_hang_hoa_format.cell(row=row_idx, column=ma_hang_col, value=ma_sp_value)
            trang_hang_hoa_format.cell(row=row_idx, column=ten_hang_col, value=ten_sp_value)
            trang_hang_hoa_format.cell(row=row_idx, column=UOM_col, value=UOM1_value)


        # Save the workbook
        file_format.save(formated)
        print(f"Successfully copied data from 'Mã SP' to 'Mã Hàng'")
        
    except Exception as e:
        print(f"Error: {str(e)}")
        raise

def normalize(text):
    normalized_name = text.lower()
    normalized_name = normalized_name.replace('à', 'a').replace('á', 'a').replace('ạ', 'a').replace('ả', 'a').replace('ã', 'a')
    normalized_name = normalized_name.replace('ằ', 'a').replace('ắ', 'a').replace('ặ', 'a').replace('ẳ', 'a').replace('ẵ', 'a')
    normalized_name = normalized_name.replace('ầ', 'a').replace('ấ', 'a').replace('ậ', 'a').replace('ẩ', 'a').replace('ẫ', 'a')
    normalized_name = normalized_name.replace('è', 'e').replace('é', 'e').replace('ẹ', 'e').replace('ẻ', 'e').replace('ẽ', 'e')
    normalized_name = normalized_name.replace('ề', 'e').replace('ế', 'e').replace('ệ', 'e').replace('ể', 'e').replace('ễ', 'e')
    normalized_name = normalized_name.replace('ì', 'i').replace('í', 'i').replace('ị', 'i').replace('ỉ', 'i').replace('ĩ', 'i')
    normalized_name = normalized_name.replace('ò', 'o').replace('ó', 'o').replace('ọ', 'o').replace('ỏ', 'o').replace('õ', 'o')
    normalized_name = normalized_name.replace('ồ', 'o').replace('ố', 'o').replace('ộ', 'o').replace('ổ', 'o').replace('ỗ', 'o')
    normalized_name = normalized_name.replace('ờ', 'o').replace('ớ', 'o').replace('ợ', 'o').replace('ở', 'o').replace('ỡ', 'o')
    normalized_name = normalized_name.replace('ù', 'u').replace('ú', 'u').replace('ụ', 'u').replace('ủ', 'u').replace('ũ', 'u')
    normalized_name = normalized_name.replace('ừ', 'u').replace('ứ', 'u').replace('ự', 'u').replace('ử', 'u').replace('ữ', 'u')
    normalized_name = normalized_name.replace('ỳ', 'y').replace('ý', 'y').replace('ỵ', 'y').replace('ỷ', 'y').replace('ỹ', 'y')
    normalized_name = normalized_name.replace('đ', 'd')
    return normalized_name

#Call
filepath = 'Raw.xlsx'
formated = 'Form.xlsx'
fill_page_hang_hoa(filepath, formated)

"""
def create_page (page_name, file_path):
    try:
        wb.create_sheet(page_name)
        wb.save(filename=file_path)
    except Exception as e:
        print(f"Error reading Excel file: {e}")
        return None
    
def fill_page_header(sheet_name, filepath, json_config):
    try:
        # Load the workbook
        wb = load_workbook(filepath)
        
        # Create or select the sheet
        if sheet_name in wb.sheetnames:
            ws = wb[sheet_name]
            ws.delete_rows(1, ws.max_row)  # Clear existing data but keep sheet
        else:
            ws = wb.create_sheet(sheet_name)
        
        # Load header configuration from JSON
        if isinstance(json_config, str):
            with open(json_config, 'r', encoding='utf-8') as f:
                config = json.load(f)
        else:
            config = json_config
            
        headers = config['headers']
        
        # Define color mapping
        color_map = {
            'green': '92D050',
            'red': 'FF0000',
        }
        
        # Write headers with formatting
        for col_num, header in enumerate(headers, 1):
            cell = ws.cell(row=1, column=col_num, value=header['name'])
            
            # Apply styling
            cell.font = Font(bold=True, name='Times New Roman')
            fill_color = color_map.get(header.get('color', 'green'), '92D050')
            cell.fill = PatternFill(start_color=fill_color, end_color=fill_color, fill_type="solid")
        
        for col_num, header in enumerate(headers, 1):
            col_letter = get_column_letter(col_num)
            header_text = header['name']
        
        # Freeze the header row
        ws.freeze_panes = "A2"
                
        # Save the changes
        wb.save(filepath)
        print(f"✅ Đã điền {len(headers)} cột vào sheet '{sheet_name}' trong file {filepath}")
        
    except Exception as e:
        print(f"❌ Lỗi khi điền dữ liệu: {str(e)}")



def find_page ():
    
    return ""

"""
# Example usage
