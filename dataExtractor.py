import pandas as pd
import os

# Folder path where the Excel files are located
folder_path = 'C:/Data Leads'

# List of columns to extract
columns_to_extract = ['COMPANY_NAME', 'FIRSTNAME', 'LASTNAME', 'JOB_TITLE', 'INDUSTRY', 'Email', 'MOBILE', 'PHONE', 'TELEPHONE', 'Clicks']
# Columns to consider for deduplication
deduplication_columns = ['COMPANY_NAME', 'FIRSTNAME', 'LASTNAME', 'JOB_TITLE', 'INDUSTRY',]

# List to store extracted data from each file
data_list = []

# Iterate over each file in the folder
for filename in os.listdir(folder_path):
    if filename.endswith('.xlsx') or filename.endswith('.xls'):
        # Construct the full file path
        file_path = os.path.join(folder_path, filename)
        
        try:
            # Read the Excel file
            df = pd.read_excel(file_path)
        
            if all(column in df.columns for column in columns_to_extract):
                # Extract the desired columns
                extracted_data:any = df[columns_to_extract]
        
                # Append the extracted data to the list
                data_list.append(extracted_data)     
            else:
                print(f"Skipped file '{filename}'. Missing one or more columns.")
        except Exception as e:
            print(f"Error occurred while processing file '{filename}': {str(e)}")


# Concatenate all data into a single DataFrame
combined_data:any = pd.concat(data_list, ignore_index=True)

# Write the combined data to a new Excel file
output_file = 'C:/Data Leads/Call List.xlsx'

try:
    combined_data.to_excel(output_file, index=False)
    print(f'Data extraction and writing completed! Output file: {output_file}')
except Exception as e:
    print(f"Error occurred while writing to the output file: {str(e)}")

# Read the Excel file
input_file = 'C:/Data Leads/Call List.xlsx'
df = pd.read_excel(input_file)

# Remove duplicate rows
df_cleaned:any = df.drop_duplicates(subset=deduplication_columns)

# Write the cleaned data to a new Excel file
output_file = 'C:/Data Leads/Call List.xlsx'

try:
    df_cleaned.to_excel(output_file, index=False)
    print(f'Data deduplication completed! Output file: {output_file}')
except Exception as e:
    print(f"Error occurred while writing to the output file: {str(e)}")


