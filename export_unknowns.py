import pandas as pd

file_path = '/Users/lakshmisk/Desktop/MLC spread sheet/spread sheet/sheet.xlsx'
output_file = 'unknown_names.txt'

excel = pd.ExcelFile(file_path)

# Show available sheets with numbers
print("Available sheets:")
for i, name in enumerate(excel.sheet_names):
    print(f"{i}: {name}")

# User picks sheet by number
sheet_index = int(input("Enter sheet number: "))
target_sheet = excel.sheet_names[sheet_index]
print(f"Processing sheet: {target_sheet}")

df = pd.read_excel(file_path, sheet_name=target_sheet)

unknowns = set()

if 'Candidate Name' in df.columns and 'Gender' in df.columns:
    unknown_rows = df[df['Gender'].astype(str).str.lower() == 'unknown']

    for val in unknown_rows['Candidate Name'].dropna():
        parts = str(val).strip().replace('.', ' ').split()
        first_name = max(parts, key=len) if parts else ''
        if first_name:
            unknowns.add(first_name)

# Write unknown names to txt (append mode so it keeps adding)
with open(output_file, 'w', encoding='utf-8') as f:
    for name in sorted(unknowns, key=lambda s: s.lower()):
        f.write(name + "\n")

print(f"✅ Added {len(unknowns)} unknown names from '{target_sheet}' to {output_file}")
