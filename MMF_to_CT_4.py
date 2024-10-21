import pandas as pd

# Load the input files
input_file = pd.ExcelFile('test stuff.xlsx')
lookup_file = pd.ExcelFile('Location Match.xlsx')

# Read the necessary sheets
report_sheet = input_file.parse('Report - Newly Created Recipes')
lookup_sheet = lookup_file.parse('Location Match')

# Initialize the output DataFrame
output_df = pd.DataFrame(columns=['Location Code (6)', 'Product # (40)', 'PLU (12)', 'Recipe Price (12,2)', 'Production Type (P/O)', 'POS Decrement (I/C)', 'Active (Y/N)', 'Recipe Location Name (80)', 'A/T Cost Item (Y/N)', 'Include in Prep Report flag (Y/N)'])

# Function to get location code from lookup sheet
def get_location_code(recipe_location_name):
    match = lookup_sheet[lookup_sheet.iloc[:, 3] == recipe_location_name]
    if not match.empty:
        return match.iloc[0, 4]
    return None

# Process each row in the report sheet
for index, row in report_sheet.iterrows():
    production_kitchen = row['Production Kitchen']
    location_serving_unit = row['Location / Serving Unit']
    daily_prep = row['Daily Prep']
    recipe_number = row['Recipe Number']
    
    # Check for exact match in lookup table
    match = lookup_sheet[(lookup_sheet.iloc[:, 2] == production_kitchen) & (lookup_sheet.iloc[:, 3] == location_serving_unit)]
    
    if not match.empty:
        # Exact match found
        for _ in range(2):
            new_row = {
                'Location Code (6)': get_location_code(production_kitchen),
                'Product # (40)': f'R{recipe_number}',
                'PLU (12)': recipe_number,
                'Recipe Price (12,2)': '',  # Assuming this value is not provided
                'Production Type (P/O)': 'P',
                'POS Decrement (I/C)': 'I' if daily_prep == 'Yes' else 'C',
                'Active (Y/N)': 'Y',  # Assuming this value is always 'Y'
                'Recipe Location Name (80)': production_kitchen,
                'A/T Cost Item (Y/N)': 'N',  # Assuming this value is always 'N'
                'Include in Prep Report flag (Y/N)': 'Y' if daily_prep == 'Yes' else 'N'
            }
            output_df = pd.concat([output_df, pd.DataFrame([new_row])], ignore_index=True)
    else:
        # No exact match found
        for _ in range(2):
            new_row = {
                'Location Code (6)': get_location_code(production_kitchen),
                'Product # (40)': f'R{recipe_number}',
                'PLU (12)': recipe_number,
                'Recipe Price (12,2)': '',  # Assuming this value is not provided
                'Production Type (P/O)': 'O',
                'POS Decrement (I/C)': 'I',
                'Active (Y/N)': 'Y',  # Assuming this value is always 'Y'
                'Recipe Location Name (80)': production_kitchen,
                'A/T Cost Item (Y/N)': 'N',  # Assuming this value is always 'N'
                'Include in Prep Report flag (Y/N)': 'N' if daily_prep == 'No' else 'Y'
            }
            output_df = pd.concat([output_df, pd.DataFrame([new_row])], ignore_index=True)
        
        # Third row with different Recipe Location Name
        match = lookup_sheet[lookup_sheet.iloc[:, 3] == location_serving_unit]
        if not match.empty:
            recipe_location_name = match.iloc[0, 2]
            new_row = {
                'Location Code (6)': get_location_code(recipe_location_name),
                'Product # (40)': f'R{recipe_number}',
                'PLU (12)': recipe_number,
                'Recipe Price (12,2)': '',  # Assuming this value is not provided
                'Production Type (P/O)': 'O',
                'POS Decrement (I/C)': 'I',
                'Active (Y/N)': 'Y',  # Assuming this value is always 'Y'
                'Recipe Location Name (80)': recipe_location_name,
                'A/T Cost Item (Y/N)': 'N',  # Assuming this value is always 'N'
                'Include in Prep Report flag (Y/N)': 'N' if daily_prep == 'No' else 'Y'
            }
            output_df = pd.concat([output_df, pd.DataFrame([new_row])], ignore_index=True)

# Save the output to a new Excel file
output_df.to_excel('text output.xlsx', index=False)