import pandas as pd

# Read the xlsx file
df = pd.read_excel('your_file.xlsx')

# Create a new column 'gender' by copying 'student_gender'
df['gender'] = df['student_gender']

# Save the updated file
df.to_excel('your_file.xlsx', index=False)

print("New column 'gender' created successfully!")