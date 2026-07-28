import os
import re
import datetime

workspace_dir = r"e:\Ask_Finance\financeapril"
exclude_dirs = {".git", "__pycache__", "finance_env", ".gemini", "IBMBOB_Demo_Assets_Mar2026"}

def fix_and_validate_date(match):
    date_str = match.group(0)
    
    # Format YYYY-MM-DD
    if re.match(r'^\d{4}-\d{2}-\d{2}$', date_str):
        year, month, day = map(int, date_str.split('-'))
        # Fix the invalid 06-31 inherited from May
        if month == 6 and day == 31:
            month = 7
            day = 31
        elif month == 6:
            month = 7
            
        try:
            # Validate
            datetime.date(year, month, day)
            return f"{year:04d}-{month:02d}-{day:02d}"
        except ValueError:
            # If invalid even after fixing (or untouched invalid), return original and let validation catch it later if it fails
            return f"{year:04d}-{month:02d}-{day:02d}"

    # Format MM/DD/YYYY
    elif re.match(r'^\d{2}/\d{2}/\d{4}$', date_str):
        parts = date_str.split('/')
        month, day, year = int(parts[0]), int(parts[1]), int(parts[2])
        if month == 6 and day == 31:
            month = 7
            day = 31
        elif month == 6:
            month = 7
            
        try:
            datetime.date(year, month, day)
            return f"{month:02d}/{day:02d}/{year:04d}"
        except ValueError:
            return f"{month:02d}/{day:02d}/{year:04d}"
            
    return date_str

def process_content(content):
    # 1. Process all YYYY-MM-DD dates using regex
    content = re.sub(r'\b\d{4}-\d{2}-\d{2}\b', fix_and_validate_date, content)
    
    # 2. Process all MM/DD/YYYY dates
    content = re.sub(r'\b\d{2}/\d{2}/\d{4}\b', fix_and_validate_date, content)
    
    # 3. Process YYYY-MM formats (e.g. 2026-06)
    content = re.sub(r'\b(2025|2026)-06\b', r'\1-07', content)
    
    # 4. Standard string replacements for non-date references
    replacements = [
        (r'\bJun 2026\b', 'Jul 2026'),
        (r'\bJun_2026\b', 'Jul_2026'),
        (r'\bJun_2025\b', 'Jul_2025'),
        (r'Jun2026', 'Jul2026'),
        (r'Jun2025', 'Jul2025'),
        (r'\bJune\b', 'July'),
        (r'\bJUNE\b', 'JULY'),
        (r'\bJun\b', 'Jul'),
        (r'\bJUN\b', 'JUL')
    ]
    
    for pattern, replacement in replacements:
        content = re.sub(pattern, replacement, content)
        
    return content

def main():
    files_to_rename = []
    
    for root, dirs, files in os.walk(workspace_dir):
        dirs[:] = [d for d in dirs if d not in exclude_dirs]
        
        for file in files:
            if file == "migrate_to_july.py" or file.endswith('.xlsm') or file.endswith('.png'):
                continue
                
            file_path = os.path.join(root, file)
            
            # Read and process content
            try:
                with open(file_path, 'r', encoding='utf-8') as f:
                    content = f.read()
            except UnicodeDecodeError:
                continue # Skip binary or weird encoding
                
            new_content = process_content(content)
            
            if new_content != content:
                with open(file_path, 'w', encoding='utf-8') as f:
                    f.write(new_content)
                print(f"Updated content in: {file_path}")
                
            # Check if file needs renaming
            new_name = file
            new_name = new_name.replace('Jun2026', 'Jul2026')
            new_name = new_name.replace('Jun2025', 'Jul2025')
            new_name = new_name.replace('june', 'july')
            new_name = new_name.replace('June', 'July')
            
            if new_name != file:
                files_to_rename.append((file_path, os.path.join(root, new_name)))

    # Rename files
    for old_path, new_path in files_to_rename:
        os.rename(old_path, new_path)
        print(f"Renamed: {os.path.basename(old_path)} -> {os.path.basename(new_path)}")
        
    print("Migration processing complete.")

if __name__ == "__main__":
    main()
