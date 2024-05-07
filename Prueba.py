import pandas as pd
import openpyxl

# Maximum number of entries per XLSX file
MAX_ENTRIES_PER_FILE = 248

# Load the complete CSV file with the updated column index
df = pd.read_csv('worldcities.csv', index_col=0)

# Group by country
groups_by_country = df.groupby('country')

# Counter for the number of processed entries
entry_counter = 0

# Counter for created files
file_counter = 0

# List to store temporary DataFrames
temp_dfs = []

# Iterate over the groups and save each one in a separate XLSX file
for country, country_data in groups_by_country:
    # Add the entries of the current country to the list of temporary DataFrames
    temp_dfs.append(country_data)
    
    # Increment the entry counter
    entry_counter += len(country_data)
    
    # If the number of entries reaches or exceeds MAX_ENTRIES_PER_FILE, concatenate the temporary DataFrames and save them to an XLSX file
    if entry_counter >= MAX_ENTRIES_PER_FILE:
        file_name = rf"C:\New folder\worldcities_{country}_{file_counter}.xlsx"
        temp_df_concatenated = pd.concat(temp_dfs)
        
        # Set 'country' as the index in the XLSX file
        temp_df_concatenated.set_index('country', inplace=True)
        
        # Save the DataFrame to an XLSX file
        temp_df_concatenated.to_excel(file_name)
        
        # Reset the list of temporary DataFrames and the entry counter
        temp_dfs = []
        entry_counter = 0
        file_counter += 1

# Save the temporary DataFrames if there are remaining entries to process
if temp_dfs:
    file_name = rf'C:\New folder\worldcities_{country}_{file_counter}.xlsx'
    temp_df_concatenated = pd.concat(temp_dfs)
    
    # Set 'country' as the index in the XLSX file
    temp_df_concatenated.set_index('country', inplace=True)
    
    # Save the DataFrame to an XLSX file
    temp_df_concatenated.to_excel(file_name)
