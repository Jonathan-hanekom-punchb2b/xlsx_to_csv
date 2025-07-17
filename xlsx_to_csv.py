import xml.etree.ElementTree as ET
import csv
import re
import os
from zipfile import ZipFile
import argparse

def column_index_to_letter(col_num):
    """Convert column number to Excel column letter (A, B, C, etc.)"""
    result = ""
    while col_num > 0:
        col_num -= 1
        result = chr(col_num % 26 + ord('A')) + result
        col_num //= 26
    return result

def excel_cell_to_coords(cell_ref):
    """Convert Excel cell reference (A1, B2, etc.) to row, col coordinates"""
    match = re.match(r'([A-Z]+)(\d+)', cell_ref)
    if not match:
        return None, None
    
    col_letters = match.group(1)
    row_num = int(match.group(2))
    
    # Convert column letters to number
    col_num = 0
    for char in col_letters:
        col_num = col_num * 26 + (ord(char) - ord('A') + 1)
    
    return row_num - 1, col_num - 1  # Convert to 0-based indexing

def parse_shared_strings(zip_file):
    """Parse sharedStrings.xml to get string values"""
    shared_strings = []
    try:
        with zip_file.open('xl/sharedStrings.xml') as f:
            root = ET.parse(f).getroot()
            for si in root.findall('.//{http://schemas.openxmlformats.org/spreadsheetml/2006/main}si'):
                t = si.find('.//{http://schemas.openxmlformats.org/spreadsheetml/2006/main}t')
                if t is not None:
                    shared_strings.append(t.text or '')
                else:
                    shared_strings.append('')
    except KeyError:
        print("No sharedStrings.xml found - file may not have text data")
    
    return shared_strings

def parse_worksheet_xml(xml_content, shared_strings):
    """Parse worksheet XML and extract cell data"""
    root = ET.fromstring(xml_content)
    
    # Dictionary to store cell data
    cells = {}
    max_row = 0
    max_col = 0
    
    # Find all cells with data
    for row in root.findall('.//{http://schemas.openxmlformats.org/spreadsheetml/2006/main}row'):
        for cell in row.findall('.//{http://schemas.openxmlformats.org/spreadsheetml/2006/main}c'):
            cell_ref = cell.get('r')
            if not cell_ref:
                continue
                
            row_idx, col_idx = excel_cell_to_coords(cell_ref)
            if row_idx is None or col_idx is None:
                continue
            
            # Get cell value
            cell_value = ""
            value_elem = cell.find('.//{http://schemas.openxmlformats.org/spreadsheetml/2006/main}v')
            
            if value_elem is not None:
                cell_type = cell.get('t')
                if cell_type == 's':  # Shared string
                    try:
                        string_index = int(value_elem.text)
                        if string_index < len(shared_strings):
                            cell_value = shared_strings[string_index]
                    except (ValueError, IndexError):
                        cell_value = value_elem.text or ''
                else:
                    cell_value = value_elem.text or ''
            
            # Only store non-empty cells
            if cell_value.strip():
                cells[(row_idx, col_idx)] = cell_value
                max_row = max(max_row, row_idx)
                max_col = max(max_col, col_idx)
    
    return cells, max_row, max_col

def cells_to_csv(cells, max_row, max_col, output_file):
    """Convert cell dictionary to CSV file"""
    with open(output_file, 'w', newline='', encoding='utf-8') as csvfile:
        writer = csv.writer(csvfile)
        
        # Write data row by row
        for row_idx in range(max_row + 1):
            row_data = []
            for col_idx in range(max_col + 1):
                cell_value = cells.get((row_idx, col_idx), '')
                row_data.append(cell_value)
            
            # Only write rows that have at least one non-empty cell
            if any(cell.strip() for cell in row_data):
                writer.writerow(row_data)

def convert_xlsx_to_csv(xlsx_file, output_dir=None):
    """Main function to convert XLSX to CSV"""
    if output_dir is None:
        output_dir = os.path.dirname(xlsx_file) or '.'
    
    # Ensure output directory exists
    os.makedirs(output_dir, exist_ok=True)
    
    base_name = os.path.splitext(os.path.basename(xlsx_file))[0]
    
    with ZipFile(xlsx_file, 'r') as zip_file:
        # Parse shared strings
        shared_strings = parse_shared_strings(zip_file)
        
        # Find all worksheet files
        worksheet_files = [f for f in zip_file.namelist() if f.startswith('xl/worksheets/sheet') and f.endswith('.xml')]
        
        if not worksheet_files:
            print("No worksheet files found in the Excel file")
            return
        
        print(f"Found {len(worksheet_files)} worksheet(s)")
        
        for worksheet_file in sorted(worksheet_files):
            # Extract sheet number from filename
            sheet_num = re.search(r'sheet(\d+)', worksheet_file)
            sheet_name = f"sheet{sheet_num.group(1)}" if sheet_num else "sheet"
            
            print(f"Processing {worksheet_file}...")
            
            # Read worksheet XML
            with zip_file.open(worksheet_file) as f:
                xml_content = f.read()
            
            # Parse worksheet and extract data
            cells, max_row, max_col = parse_worksheet_xml(xml_content, shared_strings)
            
            if not cells:
                print(f"  No data found in {sheet_name}")
                continue
            
            print(f"  Found data in {len(cells)} cells (max row: {max_row+1}, max col: {max_col+1})")
            
            # Create CSV file
            csv_filename = os.path.join(output_dir, f"{base_name}_{sheet_name}.csv")
            cells_to_csv(cells, max_row, max_col, csv_filename)
            
            print(f"  Saved to: {csv_filename}")

def main():
    parser = argparse.ArgumentParser(description='Convert XLSX files to CSV, handling corrupted files with millions of empty cells')
    parser.add_argument('xlsx_files', nargs='+', help='Path(s) to the XLSX file(s)')
    parser.add_argument('-o', '--output', help='Output directory (default: same as input file)')
    
    args = parser.parse_args()
    
    # Process each file
    successful_conversions = 0
    failed_conversions = 0
    
    for xlsx_file in args.xlsx_files:
        print(f"\n{'='*60}")
        print(f"Processing: {xlsx_file}")
        print('='*60)
        
        if not os.path.exists(xlsx_file):
            print(f"❌ Error: File {xlsx_file} not found")
            failed_conversions += 1
            continue
        
        try:
            convert_xlsx_to_csv(xlsx_file, args.output)
            print(f"✅ Successfully converted: {xlsx_file}")
            successful_conversions += 1
        except Exception as e:
            print(f"❌ Error converting {xlsx_file}: {e}")
            failed_conversions += 1
    
    # Summary
    print(f"\n{'='*60}")
    print(f"CONVERSION SUMMARY")
    print('='*60)
    print(f"✅ Successful: {successful_conversions}")
    print(f"❌ Failed: {failed_conversions}")
    print(f"📁 Total files processed: {len(args.xlsx_files)}")
    
    if successful_conversions > 0:
        print(f"\n🎉 All CSV files saved successfully!")
    if failed_conversions > 0:
        print(f"\n⚠️  {failed_conversions} file(s) had errors - check messages above")

if __name__ == "__main__":
    main()