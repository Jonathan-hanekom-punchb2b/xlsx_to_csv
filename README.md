# XLSX to CSV Converter

A simple tool to convert corrupted Excel files (.xlsx) to clean CSV files. This tool is specifically designed to handle Excel files that crash Google Sheets due to millions of empty cells or formatting issues.

## 🚨 Problem This Solves

When clients send us Excel files that:
- Crash Google Sheets when opened
- Take forever to load
- Contain millions of empty cells
- Have corrupted formatting

This tool extracts **only the actual data** and creates clean CSV files that work perfectly in Google Sheets.

## 🔧 Setup (One-time only)

### Step 1: Install Python if not already installed
- **Windows**: Download from [python.org](https://www.python.org/downloads/) and install
- **Mac**: Python is usually pre-installed, or download from [python.org](https://www.python.org/downloads/)
- **Check if installed**: Open terminal/command prompt and type `python --version`

### Step 2: Download the Tool
1. Download the `xlsx_to_csv.py` file from this repository
2. Save it to a folder on your computer (like your Desktop)

## 📖 How to Use

### Method 1: Put Files in Same Folder
1. Put your Excel file in the same folder as the script
2. Open terminal in that folder
3. Type:
   ```
   python xlsx_to_csv.py filename.xlsx
   ```

### Method 2: Type the File Path
1. Open your terminal/command prompt
2. Type:
   ```
   python xlsx_to_csv.py "C:\path\to\your\file.xlsx"
   ```
3. Replace the path with your actual file location



## 📂 What You'll Get

The tool will create clean CSV files like:
- `filename_sheet1.csv`
- `filename_sheet2.csv`
- etc.

These files will:
- ✅ Open instantly in Google Sheets
- ✅ Contain only the actual data (no empty cells)
- ✅ Maintain proper row/column structure
- ✅ Be much smaller file sizes

## 📊 Example Output

```
Found 2 worksheet(s)
Processing xl/worksheets/sheet1.xml...
  Found data in 1,247 cells (max row: 523, max col: 8)
  Saved to: client_accounts_sheet1.csv
Processing xl/worksheets/sheet2.xml...
  Found data in 892 cells (max row: 201, max col: 6)
  Saved to: client_accounts_sheet2.csv
Conversion completed successfully!
```

## 🆘 Common Issues & Solutions

### "Python is not recognized"
**Solution**: Python isn't installed or not in your PATH. Reinstall Python and check "Add to PATH" during installation.

### "File not found"
**Solution**: Check your file path. Use quotes around the path if it contains spaces:
```
python xlsx_to_csv.py "C:\My Files\data.xlsx"
```

### "Permission denied"
**Solution**: Make sure the Excel file isn't open in another program.

### Script runs but no output
**Solution**: The Excel file might be completely empty or corrupted. Check the terminal output for error messages.

## 🎯 Tips for Success

1. **Use quotes** around file paths with spaces
2. **Check the output** - the script tells you exactly what it found
3. **Test with small files first** if you're nervous
4. **Keep the original file** - this tool doesn't modify your original Excel file

## 🔄 Advanced Usage

### Save to specific folder:
```
python xlsx_to_csv.py data.xlsx -o output_folder
```

### Process multiple files:
Run the command once for each file, or simply specify multiple files as an input
```
python xlsx_to_csv.py file1.xlsx file2.xlsx file3.xlsx
```
## 📞 Need Help?

If you run into issues:
1. Copy the exact error message
2. Note what command you typed
3. Share the file type/size you're trying to convert
4. Ask for help!

