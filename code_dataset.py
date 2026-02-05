import pandas as pd
from gender_detector.gender_detector import GenderDetector

# Initialize gender detector with default settings
detector = GenderDetector()

# Path to the xlsx file
file_path = 'spread sheet/Candidate_details_with_Category.xlsx'

# Read all sheets from the xlsx file
excel_file = pd.ExcelFile(file_path)

# Create a dictionary to store updated dataframes
updated_sheets = {}

# Process each sheet
for sheet_name in excel_file.sheet_names:
    print(f"Processing sheet: {sheet_name}")
    
    # Read the sheet
    df = pd.read_excel(file_path, sheet_name=sheet_name)
    
    # Check if 'Candidate Name' column exists
    if 'Candidate Name' in df.columns:
        # Check if 'Gender' column exists
        if 'Gender' in df.columns:
            # Function to get gender from candidate name using gender-detector
            def get_gender(name):
                try:
                    # Extract first name (first word)
                    first_name = str(name).split()[0]
                    # Get gender prediction using guess() method
                    gender_result = detector.guess(first_name)
                    return gender_result if gender_result else 'unknown'
                except:
                    return 'unknown'
            
            # Find rows with 'unknown' gender
            unknown_mask = df['Gender'] == 'unknown'
            unknown_count = unknown_mask.sum()
            
            if unknown_count > 0:
                # Apply gender detection only to rows with 'unknown' gender
                df.loc[unknown_mask, 'Gender'] = df.loc[unknown_mask, 'Candidate Name'].apply(get_gender)
                
                # Count updated results
                gender_counts = df['Gender'].value_counts()
                
                print(f"  ✓ Updated {unknown_count} unknown gender records")
                for gender_type, count in gender_counts.items():
                    print(f"    - {gender_type}: {count}")
            else:
                print(f"  ✓ No unknown gender records found in {sheet_name}")
        else:
            print(f"  ✗ 'Gender' column not found in {sheet_name}")
        
        # Store the updated dataframe
        updated_sheets[sheet_name] = df
    else:
        print(f"  ✗ 'Candidate Name' column not found in {sheet_name}")
        updated_sheets[sheet_name] = df

# Write all updated sheets back to the file
with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
    for sheet_name, df in updated_sheets.items():
        df.to_excel(writer, sheet_name=sheet_name, index=False)

print("\n✓ File updated successfully with gender-detector library for unknown values!")
