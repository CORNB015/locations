
import pandas as pd
# Load the input files
main_input_file = 'test stuff.xlsx'
lookup_file = 'Location Match.xlsx'
output_file = 'text output.xlsx'
new_ingredients_file = 'New Ingredients to be built.xlsx'
# Read the sheets
report_df = pd.read_excel(main_input_file, sheet_name='Report - Newly Created Recipes')
location_match_df = pd.read_excel(lookup_file, sheet_name='Location Match')
# Strip any leading or trailing spaces from the column names
report_df.columns = report_df.columns.str.strip()
location_match_df.columns = location_match_df.columns.str.strip()
# Initialize the output DataFrames
output_df = pd.DataFrame(columns=[
    'Location Code (6)', 'Product # (40)', 'PLU (12)', 'Recipe Price (12,2)', 
    'Production Type (P/O)', 'POS Decrement (I/C)', 'Active (Y/N)', 
    'Recipe Location Name (80)', 'A/T Cost Item (Y/N)', 'Include in Prep Report flag (Y/N)'
])
new_ingredients_df = pd.DataFrame(columns=['Product number'])
# Iterate through each row in the report
for index, row in report_df.iterrows():
    production_kitchen = row['Production Kitchen']
    location_serving_unit = row['Location / Serving Unit']
    daily_prep = row['Daily Prep']
    recipe_number = row['Recipe Number']
    
    # Check for matches in the Location Match sheet using column numbers
    match = location_match_df[
        (location_match_df.iloc[:, 2] == production_kitchen) & 
        (location_match_df.iloc[:, 3] == location_serving_unit)
    ]
    
    if not match.empty:
        production_type = 'P'
        recipe_location_names = [production_kitchen, location_serving_unit]
    else:
        production_type = 'O'
        recipe_location_names = [production_kitchen, location_serving_unit, location_match_df[location_match_df.iloc[:, 3] == location_serving_unit].iloc[0, 2]]
    
    pos_decrement = 'I' if production_type == 'O' else ('I' if daily_prep == 'Yes' else 'C')
    
    # Add rows to the output DataFrame
    for i, recipe_location_name in enumerate(recipe_location_names):
        location_code = location_match_df[
            location_match_df.iloc[:, 3] == recipe_location_name
        ].iloc[0, 4]
        if i == 0:
            production_type = 'P'
            pos_decrement = 'I' if daily_prep == 'Yes' else 'C'
        else:
            production_type = 'O'
            pos_decrement = 'I'
        
        include_in_prep_report = 'N' if production_type == 'O' and pos_decrement == 'I' else ('Y' if daily_prep == 'Yes' else 'N')
        
        new_row = pd.DataFrame({
            'Location Code (6)': [location_code],
            'Product # (40)': [f'R{recipe_number}'],
            'PLU (12)': [recipe_number],
            'Production Type (P/O)': [production_type],
            'POS Decrement (I/C)': [pos_decrement],
            'Active (Y/N)': ['Y'],
            'Include in Prep Report flag (Y/N)': [include_in_prep_report],
            'Recipe Location Name (80)': [recipe_location_name]
        })
        
        output_df = pd.concat([output_df, new_row], ignore_index=True)
    
    if production_type == 'O':
        new_row = pd.DataFrame({
            'Location Code (6)': [''],
            'Product # (40)': [f'R{recipe_number}'],
            'PLU (12)': [recipe_number],
            'Production Type (P/O)': ['P'],
            'POS Decrement (I/C)': ['I' if daily_prep == 'Yes' else 'C'],
            'Active (Y/N)': ['Y'],
            'Include in Prep Report flag (Y/N)': ['Y' if daily_prep == 'Yes' else 'N'],
            'Recipe Location Name (80)': [location_match_df[location_match_df.iloc[:, 3] == location_serving_unit].iloc[0, 2]]
        })
        
        output_df = pd.concat([output_df, new_row], ignore_index=True)
    
    # Check for new ingredients
    if row['New Ingredient?'] == 'Yes':
        for col in ['New SKU Ingredient #1', 'New SKU Ingredient #2', 'New SKU Ingredient #3']:
            if pd.notna(row[col]):
                new_ingredient = row[col]
                if len(str(new_ingredient)) == 14:
                    new_ingredient = f'P{new_ingredient}'
                elif len(str(new_ingredient)) == 7:
                    new_ingredient = f'R{new_ingredient}'
                new_ingredient_row = pd.DataFrame({'Product number': [new_ingredient]})
                new_ingredients_df = pd.concat([new_ingredients_df, new_ingredient_row], ignore_index=True)
# Remove duplicate values in the new ingredients DataFrame
new_ingredients_df = new_ingredients_df.drop_duplicates()
# Save the output DataFrames to the output files
output_df.to_excel(output_file, index=False)
new_ingredients_df.to_excel(new_ingredients_file, index=False)
