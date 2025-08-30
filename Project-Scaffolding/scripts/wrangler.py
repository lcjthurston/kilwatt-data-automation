import pandas as pd

# Source Excel file
source_file = "ERCOT.xlsx"

# Destination Excel file
new_file = "ERCOT_cleaned.xlsx"

# Read the data: skip first 11 rows of notes, use row 12 as header
df = pd.read_excel(source_file, skiprows=11)

# Write the cleaned data to a brand new Excel file
df.to_excel(new_file, index=False)

print(f"Data successfully copied to {new_file}")
