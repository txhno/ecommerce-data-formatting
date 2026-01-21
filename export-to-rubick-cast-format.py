import pandas as pd
import os

def format_excel_file(input_file, template_file, output_file):
    print("Loading files...")
    
    # Load the Excel sheets
    try:
        # Load the template to get the column structure
        df_template = pd.read_excel(template_file)
        
        # Load the input data that needs formatting
        df_input = pd.read_excel(input_file)
    except Exception as e:
        print(f"Error loading files: {e}")
        return

    # Get the list of columns from the template
    target_columns = df_template.columns.tolist()
    
    print(f"Targeting {len(target_columns)} columns from the template.")

    # Reindex the input dataframe
    # This performs 3 actions:
    # 1. Keeps columns that exist in both.
    # 2. Drops columns in Input that are NOT in Template.
    # 3. Adds columns that are in Template but NOT in Input (fills with empty values), preserving the order.
    df_output = df_input.reindex(columns=target_columns)
    
    # Save the result
    df_output.to_excel(output_file, index=False)
    print(f"Success! Formatted file saved as: {output_file}")

# --- Configuration ---
# Update these names to match your actual file names
input_xlsx = "myntra cast-2026-01-21-14-09-31.xlsx" 
template_xlsx = "Myntra CAST - Batch 45 - First 4999.xlsx"
output_xlsx = "Myntra_Formatted_Output.xlsx"

if __name__ == "__main__":
    if os.path.exists(input_xlsx) and os.path.exists(template_xlsx):
        format_excel_file(input_xlsx, template_xlsx, output_xlsx)
    else:
        print("Error: One or more input files not found in the directory.")