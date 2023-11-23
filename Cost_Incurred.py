import os
import pandas as pd
from tkinter import Tk, filedialog

# Function to get the data directory using Tkinter
def get_data_directory():
    root = Tk()
    root.withdraw()  # Hide the main window

    # Open the dialog box to select the directory
    data_directory = filedialog.askdirectory(title="Select data directory")
    return data_directory


# Function to get the job codes file using Tkinter
def get_job_codes_file():
    root = Tk()
    root.withdraw()  # Hide the main window

    # Open the dialog box to select the job codes file
    job_codes_file = (
        filedialog.askopenfilename(title="Select job codes file", filetypes=[("Excel Files", "*.xlsx")]))
    return job_codes_file


# Function to get the output file path using Tkinter
def get_output_file_path():
    root = Tk()
    root.withdraw()  # Hide the main window

    # Open the dialog box to select the output file path
    local_output_file_path = filedialog.asksaveasfilename(
        title="Save Result As", filetypes=[("Excel Files", "*.xlsx")], defaultextension=".xlsx")
    return local_output_file_path


# Get data directory, job codes file, and output file path
data_path = get_data_directory()
job_codes_file_path = get_job_codes_file()
output_file_path = get_output_file_path()

# List to store DataFrames for each file
dataframes = []

# List of columns to be removed
columns_to_remove = [
    "Estimated cost",
    "Projected cost",
    "Last cost",
    "% complete - cost"
]

# Iterate over each file in the data directory
for file in os.listdir(data_path):
    if file.endswith(".xlsx"):
        # Full path of the file
        file_path = os.path.join(data_path, file)

        # Read the Excel file and remove unwanted columns
        df = pd.read_excel(file_path)
        df = df.drop(columns=columns_to_remove, errors='ignore')

        # Filter rows starting with 'C' in the 'Cost type' column
        df = df[df['Cost type'].str.startswith('C', na=False)]

        # Add the DataFrame to the list
        dataframes.append(df)

# Concatenate all DataFrames into one, sorting by the 'Job' and 'Open commitments' columns
result = pd.concat(dataframes, axis=0, ignore_index=True, sort=False)

# Convert the 'Job' column to a numeric format
result['Job'] = pd.to_numeric(result['Job'], errors='coerce')

# Sort the DataFrame by the 'Job' and 'Open commitments' columns
result = result.sort_values(by=['Job', 'Open commitments'])

# Read the job codes file
job_codes_df = pd.read_excel(job_codes_file_path)

# Merge the DataFrames based on the 'Job' column
result = pd.merge(result, job_codes_df, on='Job', how='left')

# Create the new 'Cost Incurred' column and add a checkmark if the sum is greater than 0
result['Cost Incurred'] = (result['JTD cost'] + result['Open commitments'] > 0)

# Remove rows where 'Cost Incurred' is False
result = result[result['Cost Incurred']]

# Reorganize the columns
column_order = (['Job', 'Project Name', 'Cost Incurred'] +
                [col for col in result.columns if col not in ['Job', 'Project Name', 'Cost Incurred']])
result = result[column_order]

# Save each Job to a separate sheet in the Excel file
with pd.ExcelWriter(output_file_path, engine='xlsxwriter') as writer:
    for job, df_job in result.groupby('Job'):
        df_job.to_excel(writer, sheet_name=f'Job_{int(job)}', index=False)

        # Automatically adjust the width of the columns
        for i, col in enumerate(df_job.columns):
            max_len = max(df_job[col].astype(str).apply(len).max(), len(col))
            writer.sheets[f'Job_{int(job)}'].set_column(i, i, max_len + 2)
