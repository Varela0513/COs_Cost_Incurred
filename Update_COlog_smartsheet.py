import pandas as pd
import smartsheet
from tkinter import Tk, filedialog, simpledialog


# Function to get Smartsheet credentials using Tkinter input dialogs
def get_smartsheet_credentials():
    tk = Tk()
    tk.withdraw()

    access_token = simpledialog.askstring("Smartsheet Access Token", "Enter your Smartsheet Access Token:")
    sheet_id = simpledialog.askstring("Smartsheet Sheet ID", "Enter your Smartsheet Sheet ID:")

    return access_token, sheet_id


# Function to get the Excel file using Tkinter file dialog
def get_excel_file():
    tk = Tk()
    tk.withdraw()

    excel_file = filedialog.askopenfilename(title="Select Excel file", filetypes=[("Excel Files", "*.xlsx")])
    return excel_file


# Function to update Smartsheet and generate a report
def update_smartsheet(access_token, sheet_id, excel_file_path):
    smartsheet_client = smartsheet.Smartsheet(access_token)

    try:
        # Load Smartsheet sheet
        smartsheet_sheet = smartsheet_client.Sheets.get_sheet(sheet_id)

        # Read the Excel file
        excel_data = pd.read_excel(excel_file_path)

        # List to store 'Cost Type' not found in Smartsheet
        not_found_cost_types = []

        # Iterate over rows in the DataFrame
        for index, row in excel_data.iterrows():
            job_name = row['Project Name']
            cost_type = row['Cost type']
            cost_incurred = row['Cost Incurred']

            # Find the corresponding row in Smartsheet based on 'Job Name'
            job_found = False
            for smartsheet_row in smartsheet_sheet.rows:
                if smartsheet_row.get_column('Job Name').display_value == job_name:
                    job_found = True
                    # Find the corresponding cell in the 'VKE COP #' column
                    for cell in smartsheet_row.cells:
                        if cell.column_id == 'VKE_COP#':
                            if str(cost_type) == cell.display_value:
                                # Update 'Cost Incurred (PA Use)'
                                cell_value = 'true' if cost_incurred else 'false'
                                smartsheet_row.set_column_value('Cost Incurred (PA Use)', cell_value)
                                break

            # If 'Job Name' not found, consider all 'Cost Type' as not found
            if not job_found:
                not_found_cost_types.append(cost_type)

        # Save the updated Smartsheet
        smartsheet_client.Sheets.update_sheet(smartsheet_sheet)

        return not_found_cost_types

    except PermissionError as e:
        print(f"Error: {e}")
        print("PermissionError: You don't have permission to access the Excel file. Check file permissions.")
        return []


# Function to generate a report
def generate_report(not_found_cost_types):
    if not_found_cost_types:
        print("The following 'Cost Types' were not found in Smartsheet:")
        for cost_type in not_found_cost_types:
            print(cost_type)
    else:
        print("All 'Cost Types' were successfully updated in Smartsheet.")


# Main execution
def main():
    access_token, sheet_id = get_smartsheet_credentials()
    excel_file_path = get_excel_file()

    not_found_cost_types = update_smartsheet(access_token, sheet_id, excel_file_path)
    generate_report(not_found_cost_types)
    print("Smartsheet updated successfully!")


if __name__ == "__main__":
    main()
