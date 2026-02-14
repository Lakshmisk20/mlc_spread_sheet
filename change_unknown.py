# NOTE : this script will only update the names if the gender is unknown

import pandas as pd
from names_dictionary import female_names, male_names, female_suffixes, male_suffixes,male_contains,female_contains

# normalize once ie converting female and male names to lower case before comparison
female_names = {n.lower() for n in female_names}
male_names   = {n.lower() for n in male_names}
female_suffixes = {s.lower() for s in female_suffixes}
male_suffixes   = {s.lower() for s in male_suffixes}
male_contains   = {s.lower() for s in male_contains}
female_contains   = {s.lower() for s in female_contains}
# Path to the xlsx file
file_path = '/Users/lakshmisk/Desktop/MLC spread sheet/spread sheet/sheet.xlsx'

excel_file = pd.ExcelFile(file_path)
updated_sheets = {}


def get_gender_custom(name):
    """Detect gender using only `names_dictionary` (no external gender libs).

    Strategy:
    1. Exact match in female_names or male_names
    2. Suffix-based heuristics
    3. Return 'unknown' if no match
    """
    try:
        parts = str(name).strip().replace('.', ' ').split()
        first_name = max(parts, key=len) if parts else ''
        first_name = first_name.lower().strip()
        if not first_name:
            return 'unknown'
        
        first_name = first_name.lower()   # 🔑 normalize once

        name_lower = first_name.lower()

        if name_lower in female_names:
            return 'female'

        if name_lower in male_names:
            return 'male'

        for suffix in female_suffixes:
            if name_lower.endswith(suffix):
                return 'female'

        for suffix in male_suffixes:
            if name_lower.endswith(suffix):
                return 'male'
            
        for contains in male_contains:
            if name.__contains__(contains):
                return 'male'
            
        for contains in female_contains:
            if name.__contains__(contains):
                return 'female'

        return 'unknown'

    except Exception:
        return 'unknown'


# Process each sheet and update only 'unknown' genders
for sheet_name in excel_file.sheet_names:
    print(f"Processing sheet: {sheet_name}")
    df = pd.read_excel(file_path, sheet_name=sheet_name)

    if 'Candidate Name' not in df.columns:
        print(f"  ✗ 'Candidate Name' column not found in {sheet_name}")
        updated_sheets[sheet_name] = df
        continue

    # Case 1: Gender column already exists
    if 'Gender' in df.columns:
        unknown_mask = df['Gender'].astype(str).str.lower() == 'unknown'
        unknown_count = unknown_mask.sum()

        if unknown_count > 0:
            df.loc[unknown_mask, 'Gender'] = df.loc[unknown_mask, 'Candidate Name'].apply(get_gender_custom)
            print(f"  ✓ Updated {unknown_count} unknown genders")
        else:
            print(f"  ✓ No unknown genders to update")

    # Case 2: Gender column does not exist
    else:
        df['Gender'] = df['Candidate Name'].apply(get_gender_custom)
        print(f"  ✓ Gender column created with {len(df)} records")

    # Count results
    gender_counts = df['Gender'].value_counts()
    for gender_type, count in gender_counts.items():
        print(f"    - {gender_type}: {count}")

    updated_sheets[sheet_name] = df



# Write all updated sheets back to the file
with pd.ExcelWriter(file_path, engine='openpyxl', mode='w') as writer:
    for sheet_name, df in updated_sheets.items():
        df.to_excel(writer, sheet_name=sheet_name, index=False)

print("\n✅ File updated successfully using only names_dictionary for gender detection")
