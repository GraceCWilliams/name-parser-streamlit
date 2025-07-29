import pandas as pd
import os
import re

# CONFIGURATION
input_folder = "mail_files"
name_column_letter = 'G'
output_file = "parsed_names.xlsx"

def column_letter_to_index(letter):
    return ord(letter.upper()) - ord('A')


class ParsedPerson:
    prefixes = ['mr', 'mrs', 'ms', 'miss', 'dr', 'prof']
    suffixes = ['jr', 'sr', 'ii', 'iii', 'iv', 'v', 'md', 'phd', 'esq']
    compound_last_prefixes = ['de', 'del', 'la', 'le', 'van', 'von', 'der', 'da', 'di', 'du', 'st', 'mac', 'bin', 'al', 'los', 'dos']

    def __init__(self, full_name, row, extract_plan_ssn=False):
        self.full_name = full_name
        self.row = row
        self.first_name = ""
        self.middle_name = ""
        self.last_name = ""
        self.zip_code = ""
        self.plan_number = ""
        self.ssn = ""
        self.parse_name()
        self.extract_zip_code()

        if extract_plan_ssn:
            self.extract_plan_and_ssn()

    def parse_name(self):
        if not isinstance(self.full_name, str):
            return
        name = self.full_name.lower()
        name = re.sub(r'[.,]', '', name)
        name = re.sub(r'\s+', ' ', name).strip()
        parts = name.split()

        if parts and parts[0] in self.prefixes:
            parts = parts[1:]
        if parts and parts[-1] in self.suffixes:
            parts = parts[:-1]

        if not parts:
            return

        self.first_name = parts[0].capitalize()

        if len(parts) == 2:
            self.last_name = parts[1].capitalize()
            return
        elif len(parts) == 1:
            return

        # Try to identify compound last names
        # Scan from end until we stop seeing known prefixes
        i = len(parts) - 1
        last_name_parts = [parts[i]]
        i -= 1
        while i > 0 and parts[i] in self.compound_last_prefixes:
            last_name_parts.insert(0, parts[i])
            i -= 1

        self.last_name = ' '.join(word.capitalize() for word in last_name_parts)

        middle_parts = parts[1:i+1]  # Everything between first and last
        self.middle_name = ' '.join(word.capitalize() for word in middle_parts)
    
    def is_valid_person(self):
        # Basic validity check:
        if not self.first_name or not self.last_name:
            return False

        # Check if name is mostly digits or includes company-like keywords
        combined_name = f"{self.first_name} {self.last_name}".lower()

        company_keywords = ['inc', 'llc', 'corp', 'co', 'company', 'group', 'corporation', 'pllc', 'llp', 'ltd']
        if any(word in combined_name for word in company_keywords):
            return False

        # Check for any numbers in the name (unlikely in real personal names)
        if re.search(r'\d', combined_name):
            return False

        return True


    def extract_zip_code(self):
        zip_candidates = []
        for col_idx in range(8, 13):  # Columns I–M
            if col_idx < len(self.row.index):
                col_name = self.row.index[col_idx]
                cell_value = self.row[col_name]

            if pd.notna(cell_value):
                matches = self._find_all_zips_in_text(cell_value)
                zip_candidates.extend(matches)

        if zip_candidates:
            self.zip_code = zip_candidates[-1]  # Take the last found ZIP


    @staticmethod
    def _find_all_zips_in_text(text):
        if pd.isna(text):
            return []
        if isinstance(text, str):
            return re.findall(r'\b\d{5}\b', text)
        elif isinstance(text, (int, float)):
            as_str = str(int(text)).zfill(5)
            return [as_str] if re.fullmatch(r'\d{5}', as_str) else []
        return []




    def process_names_and_zips(folder_path):
        name_col_idx = column_letter_to_index(name_column_letter)
        parsed_data = []

        for filename in os.listdir(folder_path):
            if filename.endswith('.xlsx') or filename.endswith('.xls'):
                full_path = os.path.join(folder_path, filename)
                df = pd.read_excel(full_path, engine='openpyxl')

                if name_col_idx < len(df.columns):
                    name_col = df.columns[name_col_idx]
                    for _, row in df.iterrows():
                        person = ParsedPerson(row[name_col], row)
                        parsed_data.append({
                            'FirstName': person.first_name,
                            'LastName': person.last_name,
                            'ZipCode': person.zip_code, 
                            'Plan#': person.plan_number, 
                            'SSN': person.ssn
                        })
                else:
                    print(f"Column {name_column_letter} not found in {filename}")

        result_df = pd.DataFrame(parsed_data)
        result_df.to_excel(output_file, index=False)
        print(f"Parsed names and ZIP codes saved to {output_file}")

    def extract_plan_and_ssn(self):
        col_idx = 5  # Column F (0-based index)
        if col_idx < len(self.row.index):
            col_name = self.row.index[col_idx]
            cell_value = self.row[col_name]

            if pd.notna(cell_value):
                text = str(cell_value)

                # Try to find SSN with dashes first
                ssn_match = re.search(r'\b\d{3}-\d{2}-\d{4}\b', text)
                if ssn_match:
                    self.ssn = ssn_match.group(0)
                else:
                    # Fallback: if undashed SSN embedded in string
                    digits = re.sub(r'\D', '', text)
                    if len(digits) >= 9:
                        self.ssn = f"{digits[-9:-6]}-{digits[-6:-4]}-{digits[-4:]}"  # format last 9 digits

                # Extract a 5-digit plan number from the beginning (if present)
                plan_match = re.search(r'\b\d{5}\b', text)
                if plan_match:
                    self.plan_number = plan_match.group(0)


def process_all_files(folder_path, extract_plan_ssn=False, communication_type="Print"):
    name_col_idx = column_letter_to_index(name_column_letter)
    parsed_data = []

    for filename in os.listdir(folder_path):
        if filename.endswith('.xlsx') or filename.endswith('.xls'):
            full_path = os.path.join(folder_path, filename)
            df = pd.read_excel(full_path, engine='openpyxl')

            if name_col_idx < len(df.columns):
                name_col = df.columns[name_col_idx]
                for _, row in df.iterrows():
                    person = ParsedPerson(row[name_col], row, extract_plan_ssn)
                    if not person.is_valid_person():
                        continue  # skip company or numeric-only names

                    parsed_data.append({
                        'FirstName': person.first_name,
                        'LastName': person.last_name,
                        'ZipCode': person.zip_code,
                        'Plan#': person.plan_number,
                        'SSN': person.ssn, 
                        'CommunicationType': communication_type
                    })
            else:
                print(f"Column {name_column_letter} not found in {filename}")

    return parsed_data


if __name__ == "__main__":
    all_data = []

    # Mail files: names and ZIPs only
    all_data += process_all_files("mail_files", extract_plan_ssn=False, communication_type="Print")

    # Email files: names, ZIPs, Plan#, and SSN
    all_data += process_all_files("email_files", extract_plan_ssn=True, communication_type="Email")

    # Save everything to one Excel file
    result_df = pd.DataFrame(all_data)

    # Remove any rows where any field contains "Bad SSN" (case-insensitive)
    result_df = result_df[~result_df.apply(lambda row: row.astype(str).str.contains('Bad SSN', case=False).any(), axis=1)]

    # Drop fully duplicate rows (same values across all columns)
    result_df = result_df.drop_duplicates()

    # Write cleaned data to Excel
    result_df.to_excel("parsed_names.xlsx", index=False)
    print("All parsed data saved to parsed_names.xlsx")
