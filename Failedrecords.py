import os
import pandas as pd

# Get the current project location
project_location = os.path.dirname(__file__)

# Directory containing the CSV files (relative to the project location)
directory = os.path.join(project_location, 'May25/RA')

# Path to the Excel file with supplier names to be excluded (relative to the project location)
exclusion_path = os.path.join(project_location, 'Supplier name that we can remove.xlsx')

# Read the supplier names to be excluded
exclusion_df = pd.read_excel(exclusion_path)
vendor_names_to_remove = exclusion_df['VendorName'].tolist()

# Initialize dictionaries to store data frames by type
data_frames = {'HA': [], 'PA': [], 'RA': []}
data_frames_without_file = {'HA': [], 'PA': [], 'RA': []}

# Iterate over each file in the directory
for filename in os.listdir(directory):
    if filename.endswith('.csv'):
        # Determine the type based on the filename
        if 'HA' in filename:
            file_type = 'HA'
        elif 'PA' in filename:
            file_type = 'PA'
        elif 'RA' in filename:
            file_type = 'RA'
        else:
            continue

        # Read the CSV file into DataFrames and append to the lists
        filepath = os.path.join(directory, filename)
        df = pd.read_csv(filepath)
        df1 = pd.read_csv(filepath)
        # Add a new column with the file name
        df['File'] = filename
        data_frames[file_type].append(df)
        data_frames_without_file[file_type].append(df1)

# Function to combine data frames and remove duplicates
def combine_and_filter_dfs(dfs, dfs1, file_type):
    combined_df = pd.concat(dfs, ignore_index=True)
    combined_df1 = pd.concat(dfs1, ignore_index=True)

    # Remove duplicate rows considering all columns
    combined_df_no_duplicates = combined_df.drop_duplicates()
    combined_df_no_duplicates1 = combined_df1.drop_duplicates()

    # Select the correct column for filtering
    if file_type == 'HA':
        column_name = 'BilltoVendorName'
    else:
        column_name = 'VendorName'

    # Filter out rows from combined_df_no_duplicates1 where the column matches
    filtered_combined_df_no_duplicates1 = combined_df_no_duplicates1[~combined_df_no_duplicates1[column_name].isin(vendor_names_to_remove)]

    return combined_df_no_duplicates, combined_df_no_duplicates1, filtered_combined_df_no_duplicates1

# Process each type and create corresponding Excel files
for file_type in ['HA', 'PA', 'RA']:
    if data_frames[file_type]:
        combined_df_no_duplicates, combined_df_no_duplicates1, filtered_combined_df_no_duplicates1 = combine_and_filter_dfs(data_frames[file_type], data_frames_without_file[file_type], file_type)

        # Define the path for the new Excel file (relative to the project location)
        output_path = os.path.join(project_location, f'combined_all_with_and_without_file_names_{file_type}_new.xlsx')

        # Create an Excel writer object and write both DataFrames to separate sheets
        with pd.ExcelWriter(output_path, engine='xlsxwriter') as writer:
            # Write the DataFrame with file names
            combined_df_no_duplicates.to_excel(writer, sheet_name='With File Names', index=False)
            
            # Write the DataFrame without the 'File' column
            combined_df_no_duplicates1.to_excel(writer, sheet_name='Without File Names', index=False)
            
            # Write the filtered DataFrame to a new sheet called 'New Supplier'
            filtered_combined_df_no_duplicates1.to_excel(writer, sheet_name='New Supplier', index=False)

        print(f"Excel file created successfully for {file_type} with three sheets: 'With File Names', 'Without File Names', and 'New Supplier'!")
