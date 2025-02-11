import pandas as pd
from sqlalchemy import create_engine
import openpyxl

# Read server details from an Excel file
server_details_df = pd.read_excel("server_details3.xlsx")

# Initialize an empty list to store the combined data
combined_data = []

# Iterate through servers
for index, row in server_details_df.iterrows():
    server_name = row['ServerName']
    connection_string = row['ConnectionString']
    try:
        # Define your database connection using the extracted connection string
        engine = create_engine(connection_string)
        
        # Execute the first query
        df1 = pd.read_sql("SELECT STORE FROM CONTROL WHERE REG_NUM='001'", engine)
        
        # Execute the second query
        df2 = pd.read_sql("EXEC xp_dirtree 'E:\\DB_AUTOBACKUP', 1, 1", engine)
    
        # Check if any of the DataFrames are empty
        if not df1.empty or not df2.empty:
            # Concatenate the DataFrames
            df_combined = pd.concat([df for df in [df1, df2] if not df.empty], ignore_index=True)
            # Append the combined data to the list
            combined_data.append(df_combined)
        
    except Exception as e:
        print(f"Error for server {server_name}: {str(e)}")
        continue  # Continue to the next server

# Concatenate all the data into a single DataFrame, if there's any data
if combined_data:
    final_df = pd.concat(combined_data, ignore_index=True)
    
    # Use context manager to save the final DataFrame to the sheet
    with pd.ExcelWriter("11-02-2025 QC.xlsx", engine='openpyxl') as excel_writer:
        final_df.to_excel(excel_writer, sheet_name='Combined', index=False)
    print('Data exported to output.xlsx')
else:
    print("No data to export.")
