import pandas as pd
import os
import re

# CONFIGURATION
input_folder = "mail_files"  # Folder where your mail Excel files are stored
name_column_letter = 'G'     # Column containing names
output_file = "parsed_names.xlsx"

# Known prefixes and suffixes to strip
prefixes = ['mr', 'mrs', 'ms', 'miss', 'dr', 'prof']
suffixes = ['jr', 'sr', 'ii', 'iii', 'iv', 'v', 'md', 'phd', 'esq']

def column_letter_to_index(letter):
    return ord(letter.upper()) - ord('A')

def extract_first_last(name):
    if not isinstance(name, str):
        return "", ""

    name = name.lower()
    name = re.sub(r'[.,]', '', name)
    name = re.sub(r'\s+', ' ', name).strip()
    parts = name.split()

    if parts and parts[0] in prefixes:
        parts = parts[1:]
    if parts and parts[-1] in suffixes:
        parts = parts[:-1]

    if len(parts) >= 2:
        return parts[0].capitalize(), parts[-1].capitalize()
    elif len(parts) == 1:
        return parts[0].capitalize(), ""
    else:
        return "", ""

def process_names_only(folder_path):
    name_col_idx = column_letter_to_index(name_column_letter)
    first_names, last_names = [], []

    for filename in os.listdir(folder_path):
        if filename.endswith('.xlsx') or filename.endswith('.xls'):
            full_path = os.path.join(folder_path, filename)
            df = pd.read_excel(full_path, engine='openpyxl')

            if name_col_idx < len(df.columns):
                name_col = df.columns[name_col_idx]
                for full_name in df[name_col]:
                    first, last = extract_first_last(full_name)
                    first_names.append(first)
                    last_names.append(last)
            else:
                print(f"Column {name_column_letter} not found in {filename}")

    result_df = pd.DataFrame({
        'FirstName': first_names,
        'LastName': last_names
    })

    result_df.to_excel(output_file, index=False)
    print(f"Parsed names saved to {output_file}")

if __name__ == "__main__":
    process_names_only(input_folder)
