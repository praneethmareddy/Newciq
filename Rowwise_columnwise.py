import pandas as pd
import glob
import os

# Folder containing CIQ Excel files
ciq_folder = "./ciq_files/"

# Get all Excel files in the folder
ciq_files = glob.glob(os.path.join(ciq_folder, "*.xlsx"))

# Collect all column names
all_columns = set()
for file in ciq_files:
    df = pd.read_excel(file)
    all_columns.update(df.columns)

# Define column filters for row-wise and column-wise formats
row_wise_columns = {col for col in all_columns if not any(x in col.lower() for x in ["cellid1", "cellid2", "celltype1", "celltype2"])}
column_wise_columns = {col for col in all_columns if not any(x in col.lower() for x in ["cellid", "celltype"])}

# Create DataFrames with only column names
row_wise_df = pd.DataFrame(columns=sorted(row_wise_columns))
column_wise_df = pd.DataFrame(columns=sorted(column_wise_columns))

# Save the master templates for row-wise and column-wise formats
row_wise_path = os.path.join(ciq_folder, "master_row_wise_ciq.xlsx")
column_wise_path = os.path.join(ciq_folder, "master_column_wise_ciq.xlsx")

row_wise_df.to_excel(row_wise_path, index=False)
column_wise_df.to_excel(column_wise_path, index=False)

print(f"Row-wise Master CIQ saved to {row_wise_path}")
print(f"Column-wise Master CIQ saved to {column_wise_path}")
