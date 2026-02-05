import pandas as pd
from gender_detector.gender_detector import GenderDetector

# Initialize gender detector (world works better for Indian names)
detector = GenderDetector('world')

# Your exact file path
file_path = '/Users/lakshmisk/Desktop/MLC spread sheet/spread sheet/sheet.xlsx'

excel_file = pd.ExcelFile(file_path)
updated_sheets = {}

def normalize_gender(g):
    if not g:
        return 'Unknown'
    g = g.lower()
    if 'male' in g:
        return 'Male'
    elif 'female' in g:
        return 'Female'
    else:
        return 'Unknown'

cache = {}

for sheet_name in excel_file.sheet_names:
    print(f"Processing sheet: {sheet_name}")
    df = pd.read_excel(file_path, sheet_name=sheet_name)

    if 'Candidate Name' in df.columns:

        def get_gender(name):
            try:
                first_name = str(name).strip().split()[0].lower()

                if first_name in cache:
                    return cache[first_name]

                g = detector.get_gender(first_name)
                result = normalize_gender(g)
                cache[first_name] = result
                return result
            except:
                return 'Unknown'

        df['Gender'] = df['Candidate Name'].apply(get_gender)
        print("  ✓ Gender column created")
        print(df['Gender'].value_counts())

    else:
        print(f"  ✗ 'Candidate Name' column not found in {sheet_name}")

    updated_sheets[sheet_name] = df

# Overwrite the same file
with pd.ExcelWriter(file_path, engine='openpyxl', mode='w') as writer:
    for sheet_name, df in updated_sheets.items():
        df.to_excel(writer, sheet_name=sheet_name, index=False)

print("\n✅ File updated successfully with new Gender column!")
print("🧠 Unique first names processed:", len(cache))
