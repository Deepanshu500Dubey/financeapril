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
        if month == 7:
            month = 8

        try:
            # Validate
            datetime.date(year, month, day)
            return f"{year:04d}-{month:02d}-{day:02d}"
        except ValueError:
            # If invalid even after shifting, return as-is and let validation catch it later
            return f"{year:04d}-{month:02d}-{day:02d}"

    # Format MM/DD/YYYY
    elif re.match(r'^\d{2}/\d{2}/\d{4}$', date_str):
        parts = date_str.split('/')
        month, day, year = int(parts[0]), int(parts[1]), int(parts[2])
        if month == 7:
            month = 8

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

    # 3. Process YYYY-MM formats (e.g. 2026-07)
    content = re.sub(r'\b(2025|2026)-07\b', r'\1-08', content)

    # 4. Standard string replacements for non-date references
    replacements = [
        (r'\bJul 2026\b', 'Aug 2026'),
        (r'\bJul_2026\b', 'Aug_2026'),
        (r'\bJul_2025\b', 'Aug_2025'),
        (r'Jul2026', 'Aug2026'),
        (r'Jul2025', 'Aug2025'),
        (r'\bJuly\b', 'August'),
        (r'\bJULY\b', 'AUGUST'),
        (r'\bJul\b', 'Aug'),
        (r'\bJUL\b', 'AUG')
    ]

    for pattern, replacement in replacements:
        content = re.sub(pattern, replacement, content)

    return content

def main():
    files_to_rename = []

    for root, dirs, files in os.walk(workspace_dir):
        dirs[:] = [d for d in dirs if d not in exclude_dirs]

        for file in files:
            if file in ("migrate_to_july.py", "migrate_to_august.py") or file.endswith('.xlsm') or file.endswith('.png'):
                continue

            file_path = os.path.join(root, file)

            # Read and process content
            try:
                with open(file_path, 'r', encoding='utf-8') as f:
                    content = f.read()
            except UnicodeDecodeError:
                continue  # Skip binary or weird encoding

            new_content = process_content(content)

            if new_content != content:
                with open(file_path, 'w', encoding='utf-8') as f:
                    f.write(new_content)
                print(f"Updated content in: {file_path}")

            # Check if file needs renaming
            new_name = file
            new_name = new_name.replace('Jul2026', 'Aug2026')
            new_name = new_name.replace('Jul2025', 'Aug2025')
            new_name = new_name.replace('july', 'august')
            new_name = new_name.replace('July', 'August')

            if new_name != file:
                files_to_rename.append((file_path, os.path.join(root, new_name)))

    # Rename files
    for old_path, new_path in files_to_rename:
        os.rename(old_path, new_path)
        print(f"Renamed: {os.path.basename(old_path)} -> {os.path.basename(new_path)}")

    print("Migration processing complete.")

if __name__ == "__main__":
    main()
