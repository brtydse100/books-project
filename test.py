import pandas as pd
import os


def xlsx_to_csv(xlsx_file_path, output_folder):
    # Create the output folder if it doesn't exist
    os.makedirs(output_folder, exist_ok=True)

    # Read the Excel file
    xlsx = pd.ExcelFile(xlsx_file_path)

    # Get the base name of the Excel file (without extension)
    base_name = os.path.splitext(os.path.basename(xlsx_file_path))[0]

    # Iterate through all sheets
    for sheet_name in xlsx.sheet_names:
        # Read the sheet into a pandas DataFrame
        df = pd.read_excel(xlsx, sheet_name=sheet_name)

        # Create the output CSV file name
        csv_file_name = f"{base_name}_{sheet_name}.csv"
        csv_file_path = os.path.join(output_folder, csv_file_name)

        # Save the DataFrame to a CSV file
        df.to_csv(csv_file_path, index=False)

        print(f"Converted sheet '{sheet_name}' to {csv_file_path}")

    print("Conversion completed!")

# Example usage
xlsx_file_path = r"excel\2015.xlsx"
output_folder = r"csv"

xlsx_to_csv(xlsx_file_path, output_folder)



