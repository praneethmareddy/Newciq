import pandas as pd
import glob
import os

# Folder containing CIQ Excel files
ciq_folder = "./ciq_files/"

# Get all Excel files in the folder
ciq_files = glob.glob(os.path.join(ciq_folder, "*.xlsx"))

# Collect all unique column names
all_columns = set()
for file in ciq_files:
    df = pd.read_excel(file)
    all_columns.update(df.columns)

# Save unique column names to an Excel file
unique_columns_df = pd.DataFrame(columns=sorted(all_columns))
unique_columns_path = os.path.join(ciq_folder, "unique_columns.xlsx")
unique_columns_df.to_excel(unique_columns_path, index=False)

print(f"Unique column names saved to {unique_columns_path}")
